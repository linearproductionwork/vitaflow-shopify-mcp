import express from "express";
import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { StreamableHTTPServerTransport } from "@modelcontextprotocol/sdk/server/streamableHttp.js";
import { z } from "zod";
import { createRequire } from "module";

const require = createRequire(import.meta.url);
const XLSX = require("xlsx");

const {
  SHOPIFY_SHOP,
  SHOPIFY_CLIENT_ID,
  SHOPIFY_CLIENT_SECRET,
  SHOPIFY_API_VERSION = "2026-04",
  SHOPIFY_LOCATION_ID,
  MCP_AUTH_TOKEN,
  PORT = 3000
} = process.env;

for (const [key, value] of Object.entries({ SHOPIFY_SHOP, SHOPIFY_CLIENT_ID, SHOPIFY_CLIENT_SECRET, MCP_AUTH_TOKEN })) {
  if (!value) throw new Error(`Missing required environment variable: ${key}`);
}

// ── Shopify auth ──────────────────────────────────────────────────────────────

let cachedToken = null;
let cachedTokenExpiresAt = 0;

async function getShopifyAccessToken() {
  const now = Date.now();
  if (cachedToken && now < cachedTokenExpiresAt - 60_000) return cachedToken;

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
    headers: { "Content-Type": "application/json", "X-Shopify-Access-Token": token },
    body: JSON.stringify({ query, variables })
  });

  const data = await response.json();
  if (!response.ok || data.errors) throw new Error(JSON.stringify(data.errors || data, null, 2));
  return data.data;
}

// ── Shopify helpers ───────────────────────────────────────────────────────────

let cachedLocationId = SHOPIFY_LOCATION_ID || null;

async function getPrimaryLocationId() {
  if (cachedLocationId) return cachedLocationId;
  const data = await shopifyGraphql(
    `query { locations(first: 1, includeInactive: false) { nodes { id name } } }`
  );
  if (!data.locations.nodes.length) throw new Error("No active locations found in Shopify.");
  cachedLocationId = data.locations.nodes[0].id;
  return cachedLocationId;
}

async function getProductIdFromVariant(variantId) {
  const data = await shopifyGraphql(
    `query GetVariant($id: ID!) { productVariant(id: $id) { product { id } } }`,
    { id: variantId }
  );
  return data.productVariant.product.id;
}

// Replaces deprecated productVariantUpdate
async function variantsBulkUpdate(productId, variants) {
  const data = await shopifyGraphql(
    `mutation VariantsBulkUpdate($productId: ID!, $variants: [ProductVariantsBulkInput!]!) {
      productVariantsBulkUpdate(productId: $productId, variants: $variants) {
        productVariants { id price compareAtPrice taxable }
        userErrors { field message }
      }
    }`,
    { productId, variants }
  );
  return data.productVariantsBulkUpdate;
}

// Activates inventory item at location if needed, then sets absolute quantity
async function activateAndSetInventory(inventoryItemId, locationId, quantity) {
  // Check if already activated at this location
  const check = await shopifyGraphql(
    `query CheckInventoryLevel($id: ID!) {
      inventoryItem(id: $id) {
        inventoryLevels(first: 50) {
          nodes { location { id } }
        }
      }
    }`,
    { id: inventoryItemId }
  );

  const activated = check.inventoryItem.inventoryLevels.nodes.some(
    l => l.location.id === locationId
  );

  if (!activated) {
    const activateData = await shopifyGraphql(
      `mutation ActivateInventory($itemId: ID!, $updates: [InventoryBulkToggleActivationInput!]!) {
        inventoryBulkToggleActivation(inventoryItemId: $itemId, inventoryItemUpdates: $updates) {
          inventoryItem { id }
          userErrors { field message }
        }
      }`,
      { itemId: inventoryItemId, updates: [{ locationId, activate: true }] }
    );
    const activateErrors = activateData.inventoryBulkToggleActivation.userErrors;
    if (activateErrors.length) {
      throw new Error(`Inventory activation failed: ${JSON.stringify(activateErrors)}`);
    }
  }

  const data = await shopifyGraphql(
    `mutation SetInventory($input: InventorySetQuantitiesInput!) {
      inventorySetQuantities(input: $input) {
        inventoryAdjustmentGroup { id reason }
        userErrors { field message }
      }
    }`,
    {
      input: {
        name: "available",
        reason: "correction",
        quantities: [{ inventoryItemId, locationId, quantity }]
      }
    }
  );

  const errs = data.inventorySetQuantities.userErrors;
  if (errs.length) throw new Error(`inventorySetQuantities failed: ${JSON.stringify(errs)}`);

  return data.inventorySetQuantities;
}

async function getAllShopifyProducts() {
  const all = [];
  let hasNextPage = true;
  let after = null;

  while (hasNextPage) {
    const data = await shopifyGraphql(
      `query AllProducts($first: Int!, $after: String) {
        products(first: $first, after: $after) {
          pageInfo { hasNextPage endCursor }
          nodes {
            id title handle status vendor productType tags descriptionHtml
            category { id name fullName }
            variants(first: 100) {
              nodes {
                id title sku price compareAtPrice taxable
                inventoryItem { id }
                inventoryQuantity
              }
            }
          }
        }
      }`,
      { first: 100, after }
    );
    all.push(...data.products.nodes);
    hasNextPage = data.products.pageInfo.hasNextPage;
    after = data.products.pageInfo.endCursor;
  }

  return all;
}

// ── File parsing ──────────────────────────────────────────────────────────────

function parseFileContent(content, fileType) {
  const workbook = fileType === "excel"
    ? XLSX.read(Buffer.from(content, "base64"), { type: "buffer" })
    : XLSX.read(content, { type: "string" });
  const sheet = workbook.Sheets[workbook.SheetNames[0]];
  return XLSX.utils.sheet_to_json(sheet, { defval: "" });
}

function normalizeStatus(raw) {
  const s = String(raw || "").trim().toUpperCase();
  if (s === "TRUE" || s === "ACTIVE") return "ACTIVE";
  if (s === "DRAFT" || s === "FALSE") return "DRAFT";
  if (s === "ARCHIVED") return "ARCHIVED";
  return "ACTIVE";
}

function str(v) { return String(v || "").trim(); }

function groupRowsByHandle(rows) {
  const products = new Map();

  for (const row of rows) {
    const handle = str(row["Handle"]).toLowerCase();
    if (!handle) continue;

    if (!products.has(handle)) {
      const status = row["Status"]
        ? normalizeStatus(row["Status"])
        : normalizeStatus(row["Published"]);

      products.set(handle, {
        handle,
        title: str(row["Title"]),
        bodyHtml: str(row["Body (HTML)"]),
        vendor: str(row["Vendor"]),
        productType: str(row["Type"]),
        tags: str(row["Tags"]),
        status,
        variants: []
      });
    }

    const price = str(row["Variant Price"]);
    const sku = str(row["Variant SKU"]);
    if (price || sku) {
      const taxableRaw = str(row["Variant Taxable"]).toLowerCase();
      const inventoryQty = str(row["Variant Inventory Qty"]);
      products.get(handle).variants.push({
        sku,
        price,
        compareAtPrice: str(row["Variant Compare At Price"]) || null,
        taxable: taxableRaw !== "false",
        inventoryQty: inventoryQty !== "" ? Number(inventoryQty) : null,
        option1: str(row["Option1 Value"]),
        option2: str(row["Option2 Value"]),
        option3: str(row["Option3 Value"])
      });
    }
  }

  return products;
}

// ── Diff engine ───────────────────────────────────────────────────────────────

function buildDiff(fileProducts, shopifyProducts) {
  const byHandle = new Map(shopifyProducts.map(p => [p.handle.toLowerCase(), p]));
  const byTitle  = new Map(shopifyProducts.map(p => [p.title.toLowerCase(), p]));

  const toCreate = [], toUpdate = [], unchanged = [];
  const matchedIds = new Set();

  for (const [handle, fp] of fileProducts) {
    const sp = byHandle.get(handle) ?? (fp.title ? byTitle.get(fp.title.toLowerCase()) : undefined);

    if (!sp) {
      toCreate.push({ fileProduct: fp });
      continue;
    }

    matchedIds.add(sp.id);
    const changes = [];

    if (fp.title && fp.title !== sp.title)
      changes.push({ field: "title", from: sp.title, to: fp.title });
    if (fp.vendor && fp.vendor !== sp.vendor)
      changes.push({ field: "vendor", from: sp.vendor, to: fp.vendor });
    if (fp.productType && fp.productType !== sp.productType)
      changes.push({ field: "productType", from: sp.productType, to: fp.productType });
    if (fp.status !== sp.status)
      changes.push({ field: "status", from: sp.status, to: fp.status });

    const fileTags = fp.tags.split(",").map(t => t.trim()).filter(Boolean).sort().join(",");
    const shopTags = [...sp.tags].sort().join(",");
    if (fileTags && fileTags !== shopTags)
      changes.push({ field: "tags", from: shopTags, to: fileTags });

    const variantChanges = [];
    const inventoryChanges = [];

    for (const fv of fp.variants) {
      const sv = fv.sku
        ? sp.variants.nodes.find(v => v.sku === fv.sku)
        : sp.variants.nodes[0];
      if (!sv) continue;

      if (fv.price && fv.price !== sv.price)
        variantChanges.push({ variantId: sv.id, sku: fv.sku, field: "price", from: sv.price, to: fv.price });
      if (fv.compareAtPrice !== null && fv.compareAtPrice !== (sv.compareAtPrice || ""))
        variantChanges.push({ variantId: sv.id, sku: fv.sku, field: "compareAtPrice", from: sv.compareAtPrice, to: fv.compareAtPrice });
      if (fv.taxable !== sv.taxable)
        variantChanges.push({ variantId: sv.id, sku: fv.sku, field: "taxable", from: sv.taxable, to: fv.taxable });
      if (fv.inventoryQty !== null && fv.inventoryQty !== sv.inventoryQuantity && sv.inventoryItem?.id)
        inventoryChanges.push({ inventoryItemId: sv.inventoryItem.id, variantId: sv.id, sku: fv.sku, from: sv.inventoryQuantity, to: fv.inventoryQty });
    }

    if (changes.length || variantChanges.length || inventoryChanges.length) {
      toUpdate.push({ fileProduct: fp, shopifyProduct: sp, changes, variantChanges, inventoryChanges });
    } else {
      unchanged.push({ handle: sp.handle, title: sp.title });
    }
  }

  const potentialArchives = shopifyProducts
    .filter(p => !matchedIds.has(p.id) && p.status === "ACTIVE")
    .map(p => ({ id: p.id, handle: p.handle, title: p.title, status: p.status }));

  return { toCreate, toUpdate, unchanged, potentialArchives };
}

async function applyDiff(diff) {
  const log = [];
  const locationId = await getPrimaryLocationId();

  for (const { fileProduct: fp } of diff.toCreate) {
    try {
      const input = {
        title: fp.title, status: fp.status,
        ...(fp.bodyHtml && { bodyHtml: fp.bodyHtml }),
        ...(fp.vendor && { vendor: fp.vendor }),
        ...(fp.productType && { productType: fp.productType }),
        ...(fp.tags && { tags: fp.tags.split(",").map(t => t.trim()).filter(Boolean) }),
        ...(fp.variants.length && {
          variants: fp.variants.map(v => ({
            price: v.price, taxable: v.taxable,
            ...(v.sku && { sku: v.sku }),
            ...(v.compareAtPrice && { compareAtPrice: v.compareAtPrice })
          }))
        })
      };

      const data = await shopifyGraphql(
        `mutation ProductCreate($input: ProductInput!) {
          productCreate(input: $input) {
            product { id title handle }
            userErrors { field message }
          }
        }`,
        { input }
      );

      const errs = data.productCreate.userErrors;
      log.push({ action: "create", title: fp.title, status: errs.length ? "error" : "success", ...(errs.length ? { errors: errs } : { id: data.productCreate.product.id }) });
    } catch (e) {
      log.push({ action: "create", title: fp.title, status: "error", error: e.message });
    }
  }

  for (const { shopifyProduct: sp, changes, variantChanges, inventoryChanges } of diff.toUpdate) {
    if (changes.length) {
      try {
        const input = { id: sp.id };
        for (const c of changes) {
          if (c.field === "title")       input.title = c.to;
          if (c.field === "vendor")      input.vendor = c.to;
          if (c.field === "productType") input.productType = c.to;
          if (c.field === "status")      input.status = c.to;
          if (c.field === "tags")        input.tags = c.to.split(",").map(t => t.trim()).filter(Boolean);
        }
        const data = await shopifyGraphql(
          `mutation ProductUpdate($input: ProductInput!) {
            productUpdate(input: $input) {
              product { id title }
              userErrors { field message }
            }
          }`,
          { input }
        );
        const errs = data.productUpdate.userErrors;
        log.push({ action: "update", id: sp.id, title: sp.title, status: errs.length ? "error" : "success", changes, ...(errs.length ? { errors: errs } : {}) });
      } catch (e) {
        log.push({ action: "update", id: sp.id, title: sp.title, status: "error", error: e.message });
      }
    }

    if (variantChanges.length) {
      try {
        const variantsByProduct = new Map();
        for (const vc of variantChanges) {
          if (!variantsByProduct.has(sp.id)) variantsByProduct.set(sp.id, []);
          const entry = variantsByProduct.get(sp.id).find(v => v.id === vc.variantId);
          if (entry) {
            if (vc.field === "price") entry.price = vc.to;
            if (vc.field === "compareAtPrice") entry.compareAtPrice = vc.to;
            if (vc.field === "taxable") entry.taxable = vc.to;
          } else {
            variantsByProduct.get(sp.id).push({
              id: vc.variantId,
              ...(vc.field === "price" && { price: vc.to }),
              ...(vc.field === "compareAtPrice" && { compareAtPrice: vc.to }),
              ...(vc.field === "taxable" && { taxable: vc.to })
            });
          }
        }

        for (const [productId, variants] of variantsByProduct) {
          const result = await variantsBulkUpdate(productId, variants);
          const errs = result.userErrors;
          log.push({ action: "updateVariants", productId, status: errs.length ? "error" : "success", count: variants.length, ...(errs.length ? { errors: errs } : {}) });
        }
      } catch (e) {
        log.push({ action: "updateVariants", productId: sp.id, status: "error", error: e.message });
      }
    }

    for (const ic of inventoryChanges) {
      try {
        const result = await activateAndSetInventory(ic.inventoryItemId, locationId, ic.to);
        const errs = result.userErrors;
        log.push({ action: "setInventory", productTitle: sp.title, sku: ic.sku, from: ic.from, to: ic.to, status: errs.length ? "error" : "success", ...(errs.length ? { errors: errs } : {}) });
      } catch (e) {
        log.push({ action: "setInventory", productTitle: sp.title, sku: ic.sku, status: "error", error: e.message });
      }
    }
  }

  return log;
}

// ── MCP Server ────────────────────────────────────────────────────────────────

const server = new McpServer({ name: "vitaflow-shopify-mcp", version: "2.7.0" });

server.tool(
  "list_products",
  "List Shopify products with type, tags, vendor, status, variants, category, and inventory.",
  { first: z.number().min(1).max(100).default(50) },
  async ({ first }) => {
    const data = await shopifyGraphql(
      `query Products($first: Int!) {
        products(first: $first) {
          nodes {
            id title handle status vendor productType tags description
            category { id name fullName }
            variants(first: 20) {
              nodes { id title taxable sku price compareAtPrice inventoryQuantity inventoryItem { id } }
            }
          }
        }
      }`,
      { first }
    );
    return { content: [{ type: "text", text: JSON.stringify(data.products.nodes, null, 2) }] };
  }
);

server.tool(
  "get_product_by_handle",
  "Fetch a single product by handle or search by title. Use before create_product to avoid duplicates.",
  {
    handle: z.string().optional().describe("Product handle, e.g. 'protein-isolate'"),
    title: z.string().optional().describe("Product title (partial match)")
  },
  async ({ handle, title }) => {
    if (!handle && !title) throw new Error("Provide at least handle or title.");
    const q = handle ? `handle:${handle}` : `title:${title}`;
    const data = await shopifyGraphql(
      `query SearchProducts($q: String!) {
        products(first: 10, query: $q) {
          nodes {
            id title handle status vendor productType tags
            variants(first: 10) { nodes { id sku price compareAtPrice taxable inventoryQuantity inventoryItem { id } } }
          }
        }
      }`,
      { q }
    );
    return { content: [{ type: "text", text: JSON.stringify(data.products.nodes, null, 2) }] };
  }
);

server.tool(
  "get_locations",
  "List Shopify fulfillment locations. Returns the configured primary location if SHOPIFY_LOCATION_ID is set, otherwise queries Shopify.",
  {},
  async () => {
    if (SHOPIFY_LOCATION_ID) {
      return {
        content: [{
          type: "text",
          text: JSON.stringify([{ id: SHOPIFY_LOCATION_ID, name: "Primary location (from config)", isPrimary: true }], null, 2)
        }]
      };
    }
    const data = await shopifyGraphql(
      `query { locations(first: 20, includeInactive: false) { nodes { id name address { city } isPrimary } } }`
    );
    return { content: [{ type: "text", text: JSON.stringify(data.locations.nodes, null, 2) }] };
  }
);

server.tool(
  "create_product",
  "Create a new Shopify product. Use get_product_by_handle first to confirm it doesn't exist. Use only after store owner confirms.",
  {
    title: z.string().min(1),
    handle: z.string().optional(),
    bodyHtml: z.string().optional(),
    vendor: z.string().optional(),
    productType: z.string().optional(),
    tags: z.array(z.string()).optional(),
    status: z.enum(["ACTIVE", "DRAFT"]).default("DRAFT"),
    categoryId: z.string().optional().describe("gid://shopify/TaxonomyCategory/..."),
    imageSrc: z.string().optional().describe("URL of the product image"),
    variants: z.array(z.object({
      price: z.string(),
      sku: z.string().optional(),
      taxable: z.boolean().default(true),
      requiresShipping: z.boolean().default(true),
      compareAtPrice: z.string().optional()
    })).optional()
  },
  async ({ title, handle, bodyHtml, vendor, productType, tags, status, categoryId, imageSrc, variants }) => {
    const input = {
      title, status,
      ...(handle && { handle }),
      ...(bodyHtml && { bodyHtml }),
      ...(vendor && { vendor }),
      ...(productType && { productType }),
      ...(tags?.length && { tags }),
      ...(categoryId && { category: categoryId }),
      ...(imageSrc && { images: [{ src: imageSrc }] }),
      ...(variants?.length && {
        variants: variants.map(v => ({
          price: v.price, taxable: v.taxable, requiresShipping: v.requiresShipping,
          ...(v.sku && { sku: v.sku }),
          ...(v.compareAtPrice && { compareAtPrice: v.compareAtPrice })
        }))
      })
    };

    const data = await shopifyGraphql(
      `mutation ProductCreate($input: ProductInput!) {
        productCreate(input: $input) {
          product { id title handle status category { id name } }
          userErrors { field message }
        }
      }`,
      { input }
    );
    return { content: [{ type: "text", text: JSON.stringify(data.productCreate, null, 2) }] };
  }
);

server.tool(
  "update_product_details",
  "Update title, description, vendor, productType, or tags on a product. Use only after store owner confirms.",
  {
    productId: z.string().describe("gid://shopify/Product/123"),
    title: z.string().optional(),
    bodyHtml: z.string().optional(),
    vendor: z.string().optional(),
    productType: z.string().optional(),
    tags: z.array(z.string()).optional()
  },
  async ({ productId, title, bodyHtml, vendor, productType, tags }) => {
    const input = {
      id: productId,
      ...(title !== undefined && { title }),
      ...(bodyHtml !== undefined && { bodyHtml }),
      ...(vendor !== undefined && { vendor }),
      ...(productType !== undefined && { productType }),
      ...(tags !== undefined && { tags })
    };
    const data = await shopifyGraphql(
      `mutation ProductUpdate($input: ProductInput!) {
        productUpdate(input: $input) {
          product { id title vendor productType tags descriptionHtml }
          userErrors { field message }
        }
      }`,
      { input }
    );
    return { content: [{ type: "text", text: JSON.stringify(data.productUpdate, null, 2) }] };
  }
);

server.tool(
  "update_product_type",
  "Update a Shopify product type. Use only after store owner confirms.",
  { productId: z.string(), productType: z.string().min(1) },
  async ({ productId, productType }) => {
    const data = await shopifyGraphql(
      `mutation ProductUpdate($input: ProductInput!) {
        productUpdate(input: $input) {
          product { id title productType }
          userErrors { field message }
        }
      }`,
      { input: { id: productId, productType } }
    );
    return { content: [{ type: "text", text: JSON.stringify(data.productUpdate, null, 2) }] };
  }
);

server.tool(
  "update_product_tags",
  "Replace Shopify product tags. Use only after store owner confirms.",
  { productId: z.string(), tags: z.array(z.string()).default([]) },
  async ({ productId, tags }) => {
    const data = await shopifyGraphql(
      `mutation ProductUpdate($input: ProductInput!) {
        productUpdate(input: $input) {
          product { id title tags }
          userErrors { field message }
        }
      }`,
      { input: { id: productId, tags } }
    );
    return { content: [{ type: "text", text: JSON.stringify(data.productUpdate, null, 2) }] };
  }
);

server.tool(
  "update_product_status",
  "Change a product status to ACTIVE, DRAFT, or ARCHIVED. Use only after store owner confirms.",
  {
    productId: z.string().describe("gid://shopify/Product/123"),
    status: z.enum(["ACTIVE", "DRAFT", "ARCHIVED"])
  },
  async ({ productId, status }) => {
    const data = await shopifyGraphql(
      `mutation ProductUpdate($input: ProductInput!) {
        productUpdate(input: $input) {
          product { id title status }
          userErrors { field message }
        }
      }`,
      { input: { id: productId, status } }
    );
    return { content: [{ type: "text", text: JSON.stringify(data.productUpdate, null, 2) }] };
  }
);

server.tool(
  "archive_product",
  "Archive a product (sets status ARCHIVED). Never deletes. Use only after store owner confirms.",
  { productId: z.string().describe("gid://shopify/Product/123") },
  async ({ productId }) => {
    const data = await shopifyGraphql(
      `mutation ProductUpdate($input: ProductInput!) {
        productUpdate(input: $input) {
          product { id title status }
          userErrors { field message }
        }
      }`,
      { input: { id: productId, status: "ARCHIVED" } }
    );
    return { content: [{ type: "text", text: JSON.stringify(data.productUpdate, null, 2) }] };
  }
);

server.tool(
  "update_variant_price",
  "Update price and/or compareAtPrice on a product variant. Uses productVariantsBulkUpdate. Use only after store owner confirms.",
  {
    variantId: z.string().describe("gid://shopify/ProductVariant/123"),
    price: z.string().optional(),
    compareAtPrice: z.string().optional()
  },
  async ({ variantId, price, compareAtPrice }) => {
    const productId = await getProductIdFromVariant(variantId);
    const variant = {
      id: variantId,
      ...(price !== undefined && { price }),
      ...(compareAtPrice !== undefined && { compareAtPrice })
    };
    const result = await variantsBulkUpdate(productId, [variant]);
    return { content: [{ type: "text", text: JSON.stringify(result, null, 2) }] };
  }
);

server.tool(
  "set_variant_taxable",
  "Set taxable true/false on a product variant. Uses productVariantsBulkUpdate. Use only after store owner confirms.",
  { variantId: z.string(), taxable: z.boolean() },
  async ({ variantId, taxable }) => {
    const productId = await getProductIdFromVariant(variantId);
    const result = await variantsBulkUpdate(productId, [{ id: variantId, taxable }]);
    return { content: [{ type: "text", text: JSON.stringify(result, null, 2) }] };
  }
);

server.tool(
  "set_inventory",
  "Set available stock quantity for a variant at a location. Activates the inventory item at the location automatically if needed. Use get_locations to find locationId, and list_products to find inventoryItem.id. Use only after store owner confirms.",
  {
    inventoryItemId: z.string().describe("gid://shopify/InventoryItem/... — get it from list_products or get_product_by_handle"),
    locationId: z.string().describe("gid://shopify/Location/... — get it from get_locations"),
    quantity: z.number().int().min(0)
  },
  async ({ inventoryItemId, locationId, quantity }) => {
    const result = await activateAndSetInventory(inventoryItemId, locationId, quantity);
    return { content: [{ type: "text", text: JSON.stringify(result, null, 2) }] };
  }
);

server.tool(
  "search_product_categories",
  "Search Shopify taxonomy categories by keyword. Returns IDs to use with update_product_category or create_product.",
  { query: z.string().min(1) },
  async ({ query }) => {
    const data = await shopifyGraphql(
      `query SearchCategories($query: String!) {
        taxonomy {
          categories(first: 20, query: $query) {
            nodes { id name fullName isLeaf }
          }
        }
      }`,
      { query }
    );
    return { content: [{ type: "text", text: JSON.stringify(data.taxonomy.categories.nodes, null, 2) }] };
  }
);

server.tool(
  "update_product_category",
  "Update a Shopify product taxonomy category. Use only after store owner confirms.",
  {
    productId: z.string().describe("gid://shopify/Product/123"),
    categoryId: z.string().describe("gid://shopify/TaxonomyCategory/...")
  },
  async ({ productId, categoryId }) => {
    const data = await shopifyGraphql(
      `mutation UpdateProductCategory($product: ProductUpdateInput!) {
        productUpdate(product: $product) {
          product { id title category { id name fullName } }
          userErrors { field message }
        }
      }`,
      { product: { id: productId, category: categoryId } }
    );
    return { content: [{ type: "text", text: JSON.stringify(data.productUpdate, null, 2) }] };
  }
);

server.tool(
  "sync_products_from_excel",
  [
    "Compare a Shopify products export (CSV text or base64 Excel) against the live Shopify catalog.",
    "dryRun:true returns a full plan — products to create, update, set inventory — without changing anything.",
    "dryRun:false applies all changes. Products absent from the file are flagged as potentialArchives but never auto-archived.",
    "Products are matched by Handle first, then Title. Always confirm with store owner before dryRun:false."
  ].join(" "),
  {
    fileContent: z.string().describe("Full CSV text from Shopify products export, or base64-encoded Excel file"),
    fileType: z.enum(["csv", "excel"]).default("csv"),
    dryRun: z.boolean().default(true)
  },
  async ({ fileContent, fileType, dryRun }) => {
    const rows = parseFileContent(fileContent, fileType);
    if (!rows.length) throw new Error("No rows found. Make sure to include the header row.");
    if (!("Handle" in rows[0])) throw new Error(`Missing 'Handle' column. Found: ${Object.keys(rows[0]).join(", ")}`);

    const fileProducts = groupRowsByHandle(rows);
    const shopifyProducts = await getAllShopifyProducts();
    const diff = buildDiff(fileProducts, shopifyProducts);

    if (dryRun) {
      const summary = {
        mode: "dryRun",
        summary: {
          toCreate: diff.toCreate.length,
          toUpdate: diff.toUpdate.length,
          unchanged: diff.unchanged.length,
          potentialArchives: diff.potentialArchives.length
        },
        toCreate: diff.toCreate.map(({ fileProduct: fp }) => ({
          handle: fp.handle, title: fp.title, status: fp.status, variants: fp.variants.length
        })),
        toUpdate: diff.toUpdate.map(({ shopifyProduct: sp, changes, variantChanges, inventoryChanges }) => ({
          handle: sp.handle, title: sp.title, productChanges: changes, variantChanges, inventoryChanges
        })),
        potentialArchives: diff.potentialArchives,
        unchanged: diff.unchanged.map(u => u.handle)
      };
      return { content: [{ type: "text", text: JSON.stringify(summary, null, 2) }] };
    }

    const log = await applyDiff(diff);
    const result = {
      mode: "applied",
      applied: {
        created: log.filter(l => l.action === "create" && l.status === "success").length,
        updated: log.filter(l => l.action === "update" && l.status === "success").length,
        variantBatchesUpdated: log.filter(l => l.action === "updateVariants" && l.status === "success").length,
        inventoryUpdated: log.filter(l => l.action === "setInventory" && l.status === "success").length,
        errors: log.filter(l => l.status === "error").length
      },
      potentialArchives: diff.potentialArchives,
      log
    };
    return { content: [{ type: "text", text: JSON.stringify(result, null, 2) }] };
  }
);

// ── Express ───────────────────────────────────────────────────────────────────

function assertAuthorized(req, res, next) {
  const header = req.headers.authorization || "";
  const token = header.startsWith("Bearer ") ? header.slice(7) : "";
  if (token !== MCP_AUTH_TOKEN) { res.status(401).json({ error: "Unauthorized" }); return; }
  next();
}

const app = express();
app.use(express.json({ limit: "10mb" }));

app.get("/", (_req, res) => {
  res.json({ ok: true, name: "vitaflow-shopify-mcp", version: "2.7.0", mcpEndpoint: "/mcp" });
});

app.post("/mcp", assertAuthorized, async (req, res) => {
  try {
    const transport = new StreamableHTTPServerTransport({ sessionIdGenerator: undefined });
    res.on("close", () => transport.close());
    await server.connect(transport);
    await transport.handleRequest(req, res, req.body);
  } catch (error) {
    if (!res.headersSent) {
      res.status(500).json({ jsonrpc: "2.0", error: { code: -32603, message: error.message }, id: null });
    }
  }
});

app.listen(Number(PORT), () => {
  console.log(`VitaFlow Shopify MCP v2.7.0 listening on port ${PORT}`);
});
