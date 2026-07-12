#!/usr/bin/env node
import { startMcpServer } from './server.js';

// Check for required environment variables. When WORDPRESS_USE_PG=1 we
// route REST calls through the postgres-backed PgRestRouter, so the
// WordPress/WooCommerce HTTP credentials are not required.
const requiredEnvVars = process.env.WORDPRESS_USE_PG === '1'
    ? []
    : ['WORDPRESS_SITE_URL', 'WOOCOMMERCE_CONSUMER_KEY', 'WOOCOMMERCE_CONSUMER_SECRET'];

const missingEnvVars = requiredEnvVars.filter(envVar => !process.env[envVar]);

if (missingEnvVars.length > 0) {
    console.error('Error: Missing required environment variables:');
    missingEnvVars.forEach(envVar => {
        console.error(`  - ${envVar}`);
    });
    console.error('\nPlease set these environment variables before running the server.');
    process.exit(1);
}

// Start the MCP server
startMcpServer()
    .then(() => {
        // MCP stdio protocol reserves stdout for JSON-RPC; log to stderr instead.
        console.error('InSales MCP Server started successfully');
    })
    .catch(error => {
        console.error('Failed to start InSales MCP Server:', error);
        process.exit(1);
    });