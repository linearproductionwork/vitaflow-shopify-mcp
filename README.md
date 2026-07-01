# VitaFlow Shopify MCP

MCP server for VitaFlow. Connects ChatGPT to Shopify product data for product management, category cleanup, and tax setup.

## Tools

- `list_products`: lists products, types, tags, categories, and variants.
- `update_product_type`: changes a product type after owner confirmation.
- `update_product_tags`: replaces product tags after owner confirmation.
- `set_variant_taxable`: changes a variant taxable flag after owner confirmation.
- `search_product_categories`: searches Shopify taxonomy categories by keyword.
- `update_product_category`: assigns a taxonomy category to a product after owner confirmation.

## Render Settings

- Language: `Node`
- Branch: `main`
- Build Command: `npm install`
- Start Command: `npm start`

## Environment Variables

Add these in Render under Environment (never put secrets in GitHub):

```text
SHOPIFY_SHOP=vitaflowfl.myshopify.com
SHOPIFY_CLIENT_ID=your_client_id
SHOPIFY_CLIENT_SECRET=your_client_secret
SHOPIFY_API_VERSION=2026-04
MCP_AUTH_TOKEN=make-a-long-private-password
```

## MCP URL

After Render deploys, the MCP URL will be:

```text
https://YOUR-RENDER-SERVICE.onrender.com/mcp
```

Use Bearer token authentication with the value from `MCP_AUTH_TOKEN`.
