import express from "express";
import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { StreamableHTTPServerTransport } from "@modelcontextprotocol/sdk/server/streamableHttp.js";
import { z } from "zod";

const {
  SHOPIFY_SHOP,
  SHOPIFY_CLIENT_ID,
  SHOPIFY_CLIENT_SECRET,
  SHOPIFY_API_VERSION = "2026-04",
  MCP_AUTH_TOKEN,
  PORT = 3000
} = process.env;

const requiredEnv = {
  SHOPIFY_SHOP,
  SHOPIFY_CLIENT_ID,
  SHOPIFY_CLIENT_SECRET,
  MCP_AUTH_TOKEN
};

for (const [key, value] of Object.entries(requiredEnv)) {
  if (!value) {
    throw new Error(`Missing required environment variable: ${key}`);
  }
}

let cachedToken = null;
let cachedTokenExpiresAt = 0;

function assertAuthorized(req, res, next) {
  const header = req.headers.authorization || "";
  const token = header.startsWith("Bearer ") ? header.slice(7) : "";

  if (token !== MCP_AUTH_TOKEN) {
    res.status(401).json({ error: "Unauthorized" });
    return;
  }

  next();
}

async function getShopifyAccessToken() {
  const now = Date.now();
  if (cachedToken && now < cachedTokenExpiresAt - 60_000) {
    return cachedToken;
  }

  const response = await fetch(`https://${SHOPIFY_SHOP}/admin/oauth/access_token`, {
    method: "POST",
    headers: { "Content-Type": "application/x-www-form-urlencoded" },
    body: new URLSearchParams({
      grant_type: "client_credentials",
      client_id: SHOPIFY_CLIENT_ID,
      client_secret: SHOPIFY_CLIENT_SECRET
    })
  });

  if (!response.ok) {
    const body = await response.text();
    throw new Error(`Shopify token request failed: ${response.status} ${body}`);
  }

  const data = await response.json();
  cachedToken = data.access_token;
  cachedTokenExpiresAt = now + (data.expires_in || 86399) * 1000;
  return cachedToken;
}

async function shopifyGraphql(query, variables = {}) {
  const token = await getShopifyAccessToken();
  const response = await fetch(`https://${SHOPIFY_SHOP}/admin/api/${SHOPIFY_API_VERSION}/graphql.json`, {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
      "X-Shopify-Access-Token": token
    },
    body: JSON.stringify({ query, variables })
  });

  const data = await response.json();

  if (!response.ok || data.errors) {
    throw new Error(JSON.stringify(data.errors || data, null, 2));
  }

  return data.data;
}

const server = new McpServer({
  name: "vitaflow-shopify-mcp",
  version: "1.0.0"
});

server.tool(
  "list_products",
  "List Shopify products with product type, tags, vendor, status, variants, and category.",
  {
    first: z.number().min(1).max(100).default(50)
  },
  async ({ first }) => {
    const data = await shopifyGraphql(
      `query Products($first: Int!) {
        products(first: $first) {
          nodes {
            id
            title
            handle
            status
            vendor
            productType
            tags
            description
            category {
              id
              name
              fullName
            }
            variants(first: 20) {
              nodes {
                id
                title
                taxable
                sku
                price
              }
            }
          }
        }
      }`,
      { first }
    );

    return {
      content: [
        {
          type: "text",
          text: JSON.stringify(data.products.nodes, null, 2)
        }
      ]
    };
  }
);

server.tool(
  "update_product_type",
  "Update a Shopify product type. Use only after the store owner confirms the change.",
  {
    productId: z.string().describe("Shopify product GraphQL ID"),
    productType: z.string().min(1)
  },
  async ({ productId, productType }) => {
    const data = await shopifyGraphql(
      `mutation ProductUpdate($input: ProductInput!) {
        productUpdate(input: $input) {
          product {
            id
            title
            productType
          }
          userErrors {
            field
            message
          }
        }
      }`,
      { input: { id: productId, productType } }
    );

    return {
      content: [
        {
          type: "text",
          text: JSON.stringify(data.productUpdate, null, 2)
        }
      ]
    };
  }
);

server.tool(
  "update_product_tags",
  "Replace Shopify product tags. Use only after the store owner confirms the change.",
  {
    productId: z.string().describe("Shopify product GraphQL ID"),
    tags: z.array(z.string()).default([])
  },
  async ({ productId, tags }) => {
    const data = await shopifyGraphql(
      `mutation ProductUpdate($input: ProductInput!) {
        productUpdate(input: $input) {
          product {
            id
            title
            tags
          }
          userErrors {
            field
            message
          }
        }
      }`,
      { input: { id: productId, tags } }
    );

    return {
      content: [
        {
          type: "text",
          text: JSON.stringify(data.productUpdate, null, 2)
        }
      ]
    };
  }
);

server.tool(
  "set_variant_taxable",
  "Set taxable true/false on a product variant. Use only after the store owner confirms the change.",
  {
    variantId: z.string().describe("Shopify product variant GraphQL ID"),
    taxable: z.boolean()
  },
  async ({ variantId, taxable }) => {
    const data = await shopifyGraphql(
      `mutation ProductVariantUpdate($input: ProductVariantInput!) {
        productVariantUpdate(input: $input) {
          productVariant {
            id
            title
            taxable
          }
          userErrors {
            field
            message
          }
        }
      }`,
      { input: { id: variantId, taxable } }
    );

    return {
      content: [
        {
          type: "text",
          text: JSON.stringify(data.productVariantUpdate, null, 2)
        }
      ]
    };
  }
);

server.tool(
  "search_product_categories",
  "Search Shopify product taxonomy categories by keyword. Returns category IDs and full names to use with update_product_category.",
  {
    query: z.string().min(1).describe("Search term, e.g. 'supplements', 'vitamins', 'health food'")
  },
  async ({ query }) => {
    const data = await shopifyGraphql(
      `query SearchCategories($query: String!) {
        taxonomy {
          categories(first: 20, query: $query) {
            nodes {
              id
              name
              fullName
              isLeaf
            }
          }
        }
      }`,
      { query }
    );

    return {
      content: [
        {
          type: "text",
          text: JSON.stringify(data.taxonomy.categories.nodes, null, 2)
        }
      ]
    };
  }
);

server.tool(
  "update_product_category",
  "Update a Shopify product taxonomy category. Use only after the store owner confirms the change.",
  {
    productId: z.string().describe("Shopify product GraphQL ID, e.g. gid://shopify/Product/123"),
    categoryId: z.string().describe("Shopify taxonomy category ID, e.g. gid://shopify/TaxonomyCategory/...")
  },
  async ({ productId, categoryId }) => {
    const data = await shopifyGraphql(
      `mutation UpdateProductCategory($product: ProductUpdateInput!) {
        productUpdate(product: $product) {
          product {
            id
            title
            category {
              id
              name
              fullName
            }
          }
          userErrors {
            field
            message
          }
        }
      }`,
      {
        product: {
          id: productId,
          category: categoryId
        }
      }
    );

    return {
      content: [
        {
          type: "text",
          text: JSON.stringify(data.productUpdate, null, 2)
        }
      ]
    };
  }
);

const app = express();
app.use(express.json({ limit: "1mb" }));

app.get("/", (_req, res) => {
  res.json({
    ok: true,
    name: "vitaflow-shopify-mcp",
    mcpEndpoint: "/mcp"
  });
});

app.post("/mcp", assertAuthorized, async (req, res) => {
  try {
    const transport = new StreamableHTTPServerTransport({
      sessionIdGenerator: undefined
    });

    res.on("close", () => {
      transport.close();
    });

    await server.connect(transport);
    await transport.handleRequest(req, res, req.body);
  } catch (error) {
    if (!res.headersSent) {
      res.status(500).json({
        jsonrpc: "2.0",
        error: {
          code: -32603,
          message: error.message
        },
        id: null
      });
    }
  }
});

app.listen(Number(PORT), () => {
  console.log(`VitaFlow Shopify MCP listening on port ${PORT}`);
});
