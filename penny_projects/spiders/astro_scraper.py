import scrapy
import json
from scrapy.http import Request
import xlsxwriter
from scrapy import signals
from scrapy.xlib.pydispatch import dispatcher

class AstroBrowse(scrapy.Spider):
    name = "astro_browse"
    handle_httpstatus_list = [404]

    def __init__(self,*args,**kwargs):
        self.category = kwargs.get('category','')
        self.json_file = open('astro_products_output.json','w')
        self.workbook = xlsxwriter.Workbook('astro_product_details.xlsx')
        self.worksheet = self.workbook.add_worksheet()
        self.headers_list = ["Product Id", "Product Name", "Description", "Specifications", "Categories", "Images","Related Products"]
        for i in range(0,len(self.headers_list)):
            self.worksheet.write(0,i,self.headers_list[i])
        self.values_list = []
        self.images_csv = open('product_images.csv','w')
        dispatcher.connect(self.spider_closed, signals.spider_closed)

    def spider_closed(self, spider):
        if self.values_list:
            row = 1
            col = 0
            for product_id,product_name,product_description,product_specs,product_categories,product_images,related_products_list  in (self.values_list):
                if product_images: product_images = product_images[0]
                else: product_images = ''
                data = {'ProductId': product_id , 'ProductName': product_name, 'ProductDescription': product_description, 'Specifications': product_specs,     'Categories': product_categories, 'ProductImages': product_images, 'RelatedProducts': related_products_list}
                self.json_file.write('%s\n'%json.dumps(data))
                self.images_csv.write('%s,%s\n'%(product_id,'#<>#'.join(product_images)))
                self.worksheet.write(row, col, product_id)
                self.worksheet.write(row, col + 1, product_name)
                self.worksheet.write(row, col+2, product_description)
                self.worksheet.write(row, col+3, ','.join(product_specs))
                self.worksheet.write(row, col+4, product_categories)
                self.worksheet.write(row, col+5, product_images)
                self.worksheet.write(row, col+6, json.dumps(related_products_list))
                row += 1

            self.workbook.close()
            self.images_csv.close()
            self.json_file.flush()
            self.json_file.close()


    def start_requests(self):
        if self.category:
            category_url = 'https://www.astrotools.com/product-category/%s'%(self.category.lower().replace(' ','_'))
            yield Request(category_url,callback=self.get_products_from_category)
        else:
            category_url = 'https://www.astrotools.com/product-category/'
            yield Request(category_url,callback=self.get_categories)

    def add_domain(self,url):
        if not 'www.astrotools.com' in url:
            url = 'https://www.astrotools.com' + url.replace('https:','')
        return url

    def get_categories(self,response):
        category_links = response.xpath('//div[@class="et_pb_text_inner"]//h3[contains(.,"Main Categories")]//..//following-sibling::ul//li//a//@href').extract()
        for category_url in category_links:
            category_url  = self.add_domain(category_url)
            yield Request(category_url,callback=self.get_products_from_category)

    def get_products_from_category(self,response):
        product_nodes = response.xpath('//ul[@class="products columns-4"]//li')
        for prod_node in product_nodes:
            prod_link = ''.join(prod_node.xpath('.//a//@href').extract())
            prod_link = self.add_domain(prod_link)
            product_number = ''.join(prod_node.xpath('.//a//span[@class="product_item_number"]//text()').extract()).replace('Item #','').strip()
            yield Request(prod_link,callback=self.get_product_details)

        page_navigation_link = ''.join(response.xpath('//li//a[@class="next page-numbers"]//@href').extract())
        if page_navigation_link :
            navg_url = self.add_domain(page_navigation_link )
            yield Request(navg_url,callback=self.get_products_from_category)

    def get_product_details(self,response):
        product_id = ''.join(response.xpath('//div[@class="et_pb_row et_pb_row_1_tb_body"]//div[@class="et_pb_text_inner"]//text()').extract()).replace('Item # ','').strip()
        product_name = ''.join(response.xpath('//div[@class="et_pb_module_inner"]//h1//text()').extract()).strip()
        product_description = ''.join(response.xpath('//meta[@name="description"]//@content').extract()).strip()
        product_specs = response.xpath('//div[@class="et_pb_all_tabs"]//div//div[@class="et_pb_tab_content"]//ul//li//span//text()').extract()
        product_categories = ','.join(response.xpath('//div[@class="product_meta"]//span[@class="posted_in"]//a//text()').extract())
        product_images = response.xpath('//div[contains(@class,"et_pb_module et_pb_wc_images")]//div//@data-thumb').extract()
        related_products = response.xpath('//section[@class="related products"]//ul//li')
        related_products_list = []
        for node in related_products:
            rel_dict = {}
            rel_title = ''.join(node.xpath('.//a//h2//text()').extract())
            rel_id = ''.join(node.xpath('.//a//span[@class="product_item_number"]//text()').extract())
            rel_link = ''.join(node.xpath('.//a//@href').extract())
            rel_dict = {'product_name':rel_title,'product_id':rel_id,'product_link': rel_link}
            related_products_list.append(rel_dict)
        self.values_list.append([product_id,product_name,product_description,product_specs,product_categories,product_images,related_products_list])


