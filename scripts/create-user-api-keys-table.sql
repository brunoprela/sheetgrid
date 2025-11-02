CREATE TABLE IF NOT EXISTS user_api_keys (
  user_id TEXT PRIMARY KEY NOT NULL,
  open_router_key TEXT,
  univer_mcp_key TEXT,
  created_at TIMESTAMP NOT NULL DEFAULT NOW(),
  updated_at TIMESTAMP NOT NULL DEFAULT NOW()
);

-- Create index for faster lookups
CREATE INDEX IF NOT EXISTS idx_user_api_keys_user_id ON user_api_keys(user_id);
