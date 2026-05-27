# Supabase Database Skill

## Setup
- SUPABASE_URL, SUPABASE_ANON_KEY (client)
- SUPABASE_SERVICE_KEY (server only, never expose)

## Always Use Typed Client
```typescript
import { createClient } from '@supabase/supabase-js'
import { Database } from '@/types/database'
const supabase = createClient<Database>(url, key)
```

## Error Handling
```typescript
const { data, error } = await supabase.from('users').select()
if (error) {
  console.error('Database error:', error)
  throw new Error('Failed to fetch users')
}
```

## RLS Rules
- ALWAYS enable RLS on user data tables
- Use auth.uid() to restrict to own data
- Test policies before deploying

## Common Patterns
```typescript
// Single record
const { data } = await supabase.from('profiles').select('*').eq('user_id', user.id).single()

// Insert with return
const { data } = await supabase.from('posts').insert({ title, user_id }).select().single()

// Pagination
const { data, count } = await supabase.from('items').select('*', { count: 'exact' }).range(0, 9)
```

## Don'ts
- NEVER expose service key to client
- NEVER skip error checking
- NEVER disable RLS in production
