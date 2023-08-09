# Define your item pipelines here
#
# Don't forget to add your pipeline to the ITEM_PIPELINES setting
# See: https://docs.scrapy.org/en/latest/topics/item-pipeline.html


# useful for handling different item types with a single interface
from itemadapter import ItemAdapter
import datetime
from gfinance.stock import write_into_xl


class GfinancePipeline:
    scr_file = './gfinance/data.txt'
    dest_file = 'Data.xlsx' 
    def open_spider(self, spider):
        self.fp = open(GfinancePipeline.scr_file, 'w')
        self.fp.close()
        
        self.fp = open(GfinancePipeline.scr_file, 'a')
        # Get the current date and time
        current_datetime = datetime.datetime.now()

        # Format the date and time as year/month/day hour:min AM/PM
        formatted_datetime = current_datetime.strftime("%Y/%m/%d %I:%M %p")
        self.fp.write(str(formatted_datetime))
        self.fp.write(" Eastern Time (ET)")
        self.fp.write('\n')

    def process_item(self, item, spider):
        name = item.get('name')
        price = item.get('price')
        rank = item.get('rank')
        change = item.get('change')
        p_change = item.get('p_change')

        name = "{:>30}".format(name)
        price = "{:>10}".format(price)
        rank = "{:>2}".format(rank)
        change = "{:>9}".format(change)
        # p_change = "{:>8}".format(p_change)
        p_change = str(p_change)

        if p_change.startswith("Down by"):
            p_change = "-" + p_change[8:]
        elif p_change.startswith("Up by"):
            p_change = "+" + p_change[6:]
        else:
            p_change = p_change


        self.fp.write(str(rank))
        self.fp.write('? ')
        self.fp.write(name)
        self.fp.write('? ')
        self.fp.write(price)
        self.fp.write("?  ")
        self.fp.write(change)
        self.fp.write("? ")
        self.fp.write(p_change)
        self.fp.write('\n')
        return item
    
    def close_spider(self, spider):
        self.fp.close()
        write_into_xl(GfinancePipeline.scr_file, GfinancePipeline.dest_file)