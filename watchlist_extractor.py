import asyncio
import logging
from config import CACHE_EXPIRY

class WatchlistExtractor:
    def __init__(self):
        self.cache = {}

    async def extract_watchlist(self, file_path):
        if file_path in self.cache and (asyncio.get_event_loop().time() - self.cache[file_path]['time']) < CACHE_EXPIRY:
            logging.info(f"Using cached data for {file_path}")
            return self.cache[file_path]['data']

        try:
            with open(file_path, 'r') as file:
                content = file.read().strip()
                symbols = content.split(',')
                stocks = []
                for symbol in symbols:
                    symbol = symbol.strip()
                    if symbol:
                        stocks.append({
                            'symbol': symbol,
                            'price': 0,  # These values will be updated later
                            'volume': 0,
                            'premarket_volume': 0,
                            'premarket_change_percent': 0,
                            'average_volume': 1,  # Set to 1 to avoid division by zero
                        })

            self.cache[file_path] = {'time': asyncio.get_event_loop().time(), 'data': stocks}
            logging.info(f"Successfully extracted {len(stocks)} stocks from {file_path}")
            return stocks
        except FileNotFoundError:
            logging.error(f"File not found: {file_path}")
            return []
        except Exception as e:
            logging.error(f"An error occurred while processing {file_path}: {str(e)}")
            return []

    async def extract_all_watchlists(self, file_paths):
        tasks = [self.extract_watchlist(file_path) for file_path in file_paths]
        watchlists = await asyncio.gather(*tasks)
        return [stock for watchlist in watchlists for stock in watchlist]  # Flatten the list of watchlists