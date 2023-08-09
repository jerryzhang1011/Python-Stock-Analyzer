import scrapy
from gfinance.items import GfinanceItem

class GfSpider(scrapy.Spider):
    page = 0
    page_str = ['indexes', 'most-active', 'gainers', 'losers']
    base_url = 'https://www.google.com/finance/markets/'
    name = "gf"
    allowed_domains = ["www.google.com"]
    start_urls = ["https://www.google.com/finance/markets/indexes"]
    
    def parse(self, response):        
        item = response.xpath("//div[@class='Sy70mc']//ul/li/a/div/div")
        rank = 1
        for i in item:
            stock_name = i.xpath("./div[1]/div[2]/div/text()").extract_first()
            curr_price = i.xpath("./div[2]/span/div/div/text()").extract_first()
            change = i.xpath("./div[3]/div/div/span/text()").extract_first()
            p_change = i.xpath("./div[4]/span/@aria-label").extract_first()
            data = GfinanceItem(name=stock_name, price=curr_price, change=change, p_change=p_change)
            data['rank'] = rank
            rank = rank + 1
            yield data
        
        # go to next page:
        if self.page < 3:
            self.page = self.page + 1
            ur = self.base_url + str(self.page_str[self.page])
            yield scrapy.Request(url=ur, callback=self.parse)
    

        


        
            
            
# //div[@class='Sy70mc']//ul/li/a/div/div/div[1]/div[2]/div/text() name
# //div[@class='Sy70mc']//ul/li/a/div/div/div[2]/span/div/div/text() price

