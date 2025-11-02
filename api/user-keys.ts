// Vercel Serverless Function for API key management
// This will handle server-side database access with proper security

import type { VercelRequest, VercelResponse } from '@vercel/node';
import { neon } from '@neondatabase/serverless';
import { drizzle } from 'drizzle-orm/neon-http';
import { pgTable, text, timestamp } from 'drizzle-orm/pg-core';
import { eq } from 'drizzle-orm';

const sql = neon(process.env.DATABASE_URL || '');
const db = drizzle(sql);

const userApiKeys = pgTable('user_api_keys', {
  userId: text('user_id').primaryKey().notNull(),
  openRouterKey: text('open_router_key'),
  univerMcpKey: text('univer_mcp_key'),
  createdAt: timestamp('created_at').defaultNow().notNull(),
  updatedAt: timestamp('updated_at').defaultNow().notNull(),
});

export default async function handler(req: VercelRequest, res: VercelResponse) {
  // CORS headers
  res.setHeader('Access-Control-Allow-Credentials', 'true');
  res.setHeader('Access-Control-Allow-Origin', req.headers.origin || '*');
  res.setHeader('Access-Control-Allow-Methods', 'GET,POST,OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type, Authorization');

  if (req.method === 'OPTIONS') {
    return res.status(200).end();
  }

  // Get user ID from Stack Auth token
  const authHeader = req.headers.authorization;
  if (!authHeader) {
    return res.status(401).json({ error: 'Unauthorized' });
  }

  // TODO: Verify Stack Auth token and extract user ID
  // For now, we'll accept userId from request body (insecure - needs proper auth)
  const userId = req.body?.userId || req.query?.userId;

  if (req.method === 'GET') {
    try {
      const result = await db.select().from(userApiKeys).where(eq(userApiKeys.userId, userId)).limit(1);
      return res.status(200).json(result[0] || { openRouterKey: null, univerMcpKey: null });
    } catch (error) {
      console.error('Error fetching API keys:', error);
      return res.status(500).json({ error: 'Failed to fetch API keys' });
    }
  }

  if (req.method === 'POST') {
    try {
      const { openRouterKey, univerMcpKey } = req.body;
      const now = new Date();

      const existing = await db.select().from(userApiKeys).where(eq(userApiKeys.userId, userId)).limit(1);

      if (existing[0]) {
        await db.update(userApiKeys)
          .set({
            openRouterKey: openRouterKey !== undefined ? openRouterKey : existing[0].openRouterKey,
            univerMcpKey: univerMcpKey !== undefined ? univerMcpKey : existing[0].univerMcpKey,
            updatedAt: now,
          })
          .where(eq(userApiKeys.userId, userId));
      } else {
        await db.insert(userApiKeys).values({
          userId,
          openRouterKey: openRouterKey || null,
          univerMcpKey: univerMcpKey || null,
          createdAt: now,
          updatedAt: now,
        });
      }

      return res.status(200).json({ success: true });
    } catch (error) {
      console.error('Error saving API keys:', error);
      return res.status(500).json({ error: 'Failed to save API keys' });
    }
  }

  return res.status(405).json({ error: 'Method not allowed' });
}

