import asyncio
import logging
from watchlist_extractor import WatchlistExtractor
from stock_sorter import StockSorter
from config import WATCHLIST_FILES

logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

async def main():
    extractor = WatchlistExtractor()
    sorter = StockSorter()

    try:
        # Extract stocks from all watchlist files
        watchlists = {}
        for file in WATCHLIST_FILES:
            watchlist_name = file.split('/')[-1].split('.')[0]  # Extract name from file path
            stocks = await extractor.extract_watchlist(file)
            watchlists[watchlist_name] = stocks

        logging.info(f"Extracted stocks from {len(WATCHLIST_FILES)} watchlists")

        # Sort and export stocks to Excel
        sorter.export_to_excel(watchlists, 'stock_rankings.xlsx')

        logging.info("Stocks have been sorted and exported to stock_rankings.xlsx")

        # Print a summary of the top 5 stocks overall
        all_stocks = []
        for stocks in watchlists.values():
            all_stocks.extend(stocks)
        
        sorted_stocks = sorter.sort_stocks(all_stocks)
        print("Top 5 stocks overall:")
        for i, stock in enumerate(sorted_stocks[:5], 1):
            print(f"{i}. {stock['symbol']}:")
            print(f"   Overall Rank: {stock['overall_rank']}")
            print(f"   Premarket Volume Rank: {stock['premarket_volume_rank']}")
            print(f"   Premarket Change % Rank: {stock['premarket_change_percent_rank']}")
            print(f"   Volume vs Avg Rank: {stock['volume_vs_avg_rank']}")
            print()

    except Exception as e:
        logging.error(f"An error occurred: {str(e)}")

if __name__ == "__main__":
    asyncio.run(main())