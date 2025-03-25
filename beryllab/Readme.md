Yahoo Finance (yfinance) doesn't have an officially documented API rate limit because it's an unofficial API that scrapes Yahoo Finance's website. However, users have reported experiencing limitations:
Approximately 2,000 requests per hour or about 48,000 per day
IP-based rate limiting may occur if too many requests are made too quickly
To be safe and avoid getting rate-limited, it's recommended to:
Add delays between requests (e.g., 1-2 seconds)
Implement error handling with exponential backoff
Cache results when possible