import axios, { AxiosInstance } from 'axios';
import { PgRestRouter } from './pg-rest-router.js';

// HttpLike: subset of AxiosInstance used by services. PgRestRouter implements
// the same get/post/put/delete shape, so we can swap it in transparently.
type HttpLike = Pick<AxiosInstance, 'get' | 'post' | 'put' | 'delete'>;

export class BaseService {
    protected client: HttpLike;
    protected siteUrl: string;
    protected consumerKey: string;
    protected consumerSecret: string;

    constructor() {
        this.siteUrl = process.env.WORDPRESS_SITE_URL || '';
        this.consumerKey = process.env.WOOCOMMERCE_CONSUMER_KEY || '';
        this.consumerSecret = process.env.WOOCOMMERCE_CONSUMER_SECRET || '';

        // Cowork-Bench ships InSales data in postgres (schema "wc").
        // When WORDPRESS_USE_PG=1, route REST calls through PgRestRouter
        // instead of axios — see services/pg-rest-router.ts.
        if (process.env.WORDPRESS_USE_PG === '1') {
            this.client = new PgRestRouter() as unknown as HttpLike;
            console.error('[InSales] Using PgRestRouter (postgres-backed mock)');
            return;
        }

        // 移除末尾的斜杠
        this.siteUrl = this.siteUrl.replace(/\/$/, '');

        const baseURL = `${this.siteUrl}/wp-json/wc/v3`;
        
        console.error(`[InSales] Initializing with base URL: ${baseURL}`);

        const ax: AxiosInstance = axios.create({
            baseURL,
            auth: {
                username: this.consumerKey,
                password: this.consumerSecret
            },
            headers: {
                'Content-Type': 'application/json',
            },
        });

        ax.interceptors.request.use(
            (config) => {
                console.error(`[InSales] Request: ${config.method?.toUpperCase()} ${config.baseURL}${config.url}`);
                console.error(`[InSales] Params:`, config.params);
                return config;
            },
            (error) => {
                console.error('[InSales] Request error:', error);
                return Promise.reject(error);
            }
        );

        ax.interceptors.response.use(
            (response) => {
                console.error(`[InSales] Response: ${response.status} ${response.statusText}`);
                return response;
            },
            (error) => {
                if (error.response) {
                    console.error(`[InSales] Response error: ${error.response.status} ${error.response.statusText}`);
                    console.error(`[InSales] Response URL: ${error.config?.url}`);
                    console.error(`[InSales] Response data:`, error.response.data);
                }
                return Promise.reject(error);
            }
        );

        this.client = ax;
    }

    /**
     * Convert camelCase object keys to snake_case for InSales API
     */
    protected toSnakeCase(params: any): any {
        if (!params || typeof params !== 'object') {
            return params;
        }

        const result: any = {};
        for (const key in params) {
            if (params.hasOwnProperty(key)) {
                // Convert camelCase to snake_case
                const snakeKey = key.replace(/[A-Z]/g, letter => `_${letter.toLowerCase()}`);
                result[snakeKey] = params[key];
            }
        }
        return result;
    }

    protected async handleRequest<T>(request: Promise<any>): Promise<T> {
        try {
            const response = await request;
            return response.data;
        } catch (error: any) {
            if (error.response) {
                const message = error.response.data?.message || error.response.data?.code || error.message;
                const status = error.response.status;

                if (status === 404) {
                    throw new Error(`InSales API endpoint not found. Please check if InSales is installed and REST API is enabled. URL: ${error.config?.url}`);
                } else if (status === 401) {
                    throw new Error(`InSales API authentication failed. Please check your consumer key and secret.`);
                } else if (status === 403) {
                    throw new Error(`InSales API access forbidden. Please check your API key permissions.`);
                }

                throw new Error(`InSales API error: ${message} (Status: ${status})`);
            }
            throw error;
        }
    }
}