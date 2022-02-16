import scrapy
from scrapy.http import Request
from scrapy.http.cookies import CookieJar
from bs4 import BeautifulSoup
import requests
import json
from scrapy import Selector
import xlsxwriter  
from scrapy import signals  
from scrapy.xlib.pydispatch import dispatcher

class KnipexScraper(scrapy.Spider):
    name = "knipex_browse"

    def __init__(self,*args,**kwargs):
        self.category = kwargs.get('category','')
        self.workbook = xlsxwriter.Workbook('knipex_product_details.xlsx')
        self.json_file = open('knipex_products_output.json','w')
        self.worksheet = self.workbook.add_worksheet()
        self.headers_list = ["Product Id", "Product Name", "Description", "Article No", "EAN", "Images", "Weight", "Dimensions", "Techincal Attributes"]
        for i in range(0,len(self.headers_list)):
            self.worksheet.write(0,i,self.headers_list[i])
        self.values_list = []
        self.images_csv = open('knipex_product_images.csv','w')
        dispatcher.connect(self.spider_closed, signals.spider_closed)

    def spider_closed(self, spider):
        if self.values_list:
            row = 1
            col = 0
            for product_id ,product_name ,product_description ,article_no ,ean,product_images ,weight , dimensions, technical_dict  in (self.values_list):
                if product_images:
                    product_images = product_images[0]
                else: product_images = ''
                data = {"ProductId": product_id, "ProductName": product_name, "Productdescription": product_description, 'ArticleNo': article_no, "Ean": ean, "ProductImages": product_images, "Weight":weight, "Dimensions": dimensions, "TechnicalAttributes": technical_dict}
                self.json_file.write('%s\n'%json.dumps(data))
                self.images_csv.write('%s,%s\n'%(product_id,'#<>#'.join(product_images)))
                self.worksheet.write(row, col, product_id)
                self.worksheet.write(row, col + 1, product_name)
                self.worksheet.write(row, col+2, product_description)
                self.worksheet.write(row, col+3, article_no)
                self.worksheet.write(row, col+4, ean)
                self.worksheet.write(row, col+5, product_images)
                self.worksheet.write(row, col+6, weight)
                self.worksheet.write(row, col+7, dimensions)
                self.worksheet.write(row, col+8, json.dumps(technical_dict ))
                row += 1

            self.workbook.close()
            self.images_csv.close()
            self.json_file.close()

    def start_requests(self):
        if self.category:
            category_url = 'https://www.knipex.com/products/%s'%(self.category.lower().replace(' ','-'))
            yield Request(category_url,callback=self.get_products_from_category)
        else:
            headers = {
                    'authority': 'www.knipex.com',
                    'sec-ch-ua': '" Not A;Brand";v="99", "Chromium";v="98", "Google Chrome";v="98"',
                    'sec-ch-ua-mobile': '?0',
                    'sec-ch-ua-platform': '"Windows"',
                    'upgrade-insecure-requests': '1',
                    'user-agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/98.0.4758.82 Safari/537.36',
                    'accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.9',
                    'sec-fetch-site': 'same-origin',
                    'sec-fetch-mode': 'navigate',
                    'sec-fetch-user': '?1',
                    'sec-fetch-dest': 'document',
                    'referer': 'https://www.knipex.com/',
                    'accept-language': 'en-US,en;q=0.9',
            }
            category_url = 'https://www.knipex.com/products'
            yield scrapy.Request(category_url,callback=self.get_all_categories,headers = headers)

    def add_domain(self,url):
        if not 'http' in url:
            url = 'https://www.knipex.com/products' + url
        return url


    def get_all_categories(self,response):
        category_links = response.xpath('//div[@id="block-productcategorymenu"]//ul//li//a//@href').extract()
        headers = {
                'authority': 'www.knipex.com',
                'sec-ch-ua': '" Not A;Brand";v="99", "Chromium";v="98", "Google Chrome";v="98"',
                'sec-ch-ua-mobile': '?0',
                'sec-ch-ua-platform': '"Windows"',
                'upgrade-insecure-requests': '1',
                'user-agent': 'Mozilla/5.0 AppleWebKit/537.36 (KHTML, like Gecko; compatible; Googlebot/2.1; http://www.google.com/bot.html) Chrome/W.X.Y.Z Safari/537.36',
                'accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.9',
                'sec-fetch-site': 'same-origin',
                'sec-fetch-mode': 'navigate',
                'sec-fetch-user': '?1',
                'sec-fetch-dest': 'document',
                'referer': 'https://www.knipex.com/products',
                'accept-language': 'en-US,en;q=0.9',
                'if-none-match': '"1644898437"',
        }

        for category_url in category_links:
            category_url = 'https://www.knipex.com/' + category_url #self.add_domain(category_url)
            try:
                r = requests.get(category_url)
                self.get_products_from_category(r)
            except:
                continue

    def get_products_from_category(self,response):
        sel = Selector(text=response.text)
        headers = {
            'authority': 'www.knipex.com',
            'sec-ch-ua': '"Google Chrome";v="93", " Not;A Brand";v="99", "Chromium";v="93"',
            'sec-ch-ua-mobile': '?0',
            'sec-ch-ua-platform': '"Linux"',
            'upgrade-insecure-requests': '1',
            'user-agent': 'Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/93.0.4577.63 Safari/537.36',
            'accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.9',
            'sec-fetch-site': 'same-origin',
            'sec-fetch-mode': 'navigate',
            'sec-fetch-user': '?1',
            'sec-fetch-dest': 'document',
            'accept-language': 'en-US,en;q=0.9',
        }
        product_links = sel.xpath('//div[@class="element-container"]//span[@class="field-content"]//a//@href').extract()
        for prod in list(set(product_links)):
            prod_link  = 'https://www.knipex.com' + prod 
            headers['referer'] = response.url 
            try:
                req = requests.get(prod_link)
                self.get_product_details(req)
            except: continue
        
        page_navigation = ''.join(sel.xpath('//li[@class="pager__item pager__item--next"]//a[@rel="next"]//@href').extract())
        if page_navigation:
            nav_page = response.url + page_navigation
            req = requests.get(nav_page)
            self.get_products_from_category(req)
            

    def get_product_details(self,response):
        sel = Selector(text=response.text)
        product_description = ''.join(sel.xpath('//meta[@name="description"]//@content').extract())
        product_id = response.url.split('/')[-1]
        product_name = ''.join(sel.xpath('//meta[@property="og:title"]//@content').extract())
        product_images = sel.xpath('//article[@class="ProductContainer"]//div[@class="SliderProductDetailPreview"]//div[@class="field__item"]//img//@src').extract()
        if product_images:
            product_images_list = []
            for i in product_images:
                url = 'https://www.knipex.com' + i
                product_images_list.append(url)

        technical_attr = sel.xpath('//div[contains(@class, "name-field-technical-attribute")]//following-sibling::div//div[@class="field__item"]')
        technical_dict = {}
        for i in technical_attr:
            key = ''.join(i.xpath('.//div[@class="key"]//span//text()').extract()).strip('\n').strip()
            value = ''.join(i.xpath('.//div[@class="value"]//span//text()').extract())
            if key:
                technical_dict[key] = value
        article_no = technical_dict.get('Article No.','')
        ean = technical_dict.get('EAN','')
        weight = technical_dict.get('Weight','')
        dimensions = technical_dict.get('Diemnsions','')
        self.values_list.append([product_id ,product_name ,product_description ,article_no ,ean,product_images_list ,weight , dimensions, technical_dict ])





