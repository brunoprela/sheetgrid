import { neon } from '@neondatabase/serverless';
import { drizzle } from 'drizzle-orm/neon-http';
import { pgTable, text, timestamp } from 'drizzle-orm/pg-core';
import { eq } from 'drizzle-orm';

// Use serverless Neon connection
const sql = neon(import.meta.env.VITE_DATABASE_URL || import.meta.env.DATABASE_URL || '');
export const db = drizzle(sql);

// Schema for user API keys
export const userApiKeys = pgTable('user_api_keys', {
  userId: text('user_id').primaryKey().notNull(),
  openRouterKey: text('open_router_key'),
  univerMcpKey: text('univer_mcp_key'),
  createdAt: timestamp('created_at').defaultNow().notNull(),
  updatedAt: timestamp('updated_at').defaultNow().notNull(),
});

export async function getUserApiKeys(userId: string) {
  try {
    const result = await db.select().from(userApiKeys).where(eq(userApiKeys.userId, userId)).limit(1);
    return result[0] || null;
  } catch (error) {
    console.error('Error fetching user API keys:', error);
    return null;
  }
}

export async function saveUserApiKeys(userId: string, keys: { openRouterKey?: string; univerMcpKey?: string }) {
  try {
    const now = new Date();
    const existing = await getUserApiKeys(userId);
    
    if (existing) {
      await db.update(userApiKeys)
        .set({
          openRouterKey: keys.openRouterKey !== undefined ? keys.openRouterKey : existing.openRouterKey,
          univerMcpKey: keys.univerMcpKey !== undefined ? keys.univerMcpKey : existing.univerMcpKey,
          updatedAt: now,
        })
        .where(eq(userApiKeys.userId, userId));
    } else {
      await db.insert(userApiKeys).values({
        userId,
        openRouterKey: keys.openRouterKey || null,
        univerMcpKey: keys.univerMcpKey || null,
        createdAt: now,
        updatedAt: now,
      });
    }
    return true;
  } catch (error) {
    console.error('Error saving user API keys:', error);
    throw error;
  }
}

