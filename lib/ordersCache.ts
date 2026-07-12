/**
 * Server-side in-memory cache for merged orders data.
 * Stored on `global` to survive Next.js hot reloads in development.
 * Invalidated on new imports or database clears.
 */

interface OrdersCacheEntry {
  data: any[];
  timestamp: number;
}

declare global {
  // eslint-disable-next-line no-var
  var __ordersCache: OrdersCacheEntry | null | undefined;
}

// Initialize global cache slot once
if (global.__ordersCache === undefined) {
  global.__ordersCache = null;
}

/**
 * Get the cached merged orders. Returns null if cache is empty/invalidated.
 */
export function getCachedOrders(): any[] | null {
  if (!global.__ordersCache) return null;
  return global.__ordersCache.data;
}

/**
 * Store merged orders into cache.
 */
export function setCachedOrders(data: any[]): void {
  global.__ordersCache = {
    data,
    timestamp: Date.now(),
  };
}

/**
 * Invalidate (clear) the cache. Call this after imports or database clears.
 */
export function invalidateOrdersCache(): void {
  global.__ordersCache = null;
}
