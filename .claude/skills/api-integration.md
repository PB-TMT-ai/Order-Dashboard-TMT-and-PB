# API Integration Skill

## Error Handling
```typescript
async function fetchData(endpoint: string) {
  try {
    const response = await fetch(endpoint)
    if (!response.ok) throw new Error(`HTTP ${response.status}`)
    return await response.json()
  } catch (error) {
    console.error('API Error:', error)
    throw error
  }
}
```

## Retry with Backoff
```typescript
async function fetchWithRetry(fn: () => Promise<any>, retries = 3) {
  for (let i = 0; i < retries; i++) {
    try {
      return await fn()
    } catch (e) {
      if (i === retries - 1) throw e
      await new Promise(r => setTimeout(r, 1000 * Math.pow(2, i)))
    }
  }
}
```

## Rate Limiting
- Track requests per minute
- Add delays for batch operations
- Respect Retry-After headers

## Don'ts
- NEVER hardcode API keys
- NEVER log sensitive data
- NEVER ignore rate limits
