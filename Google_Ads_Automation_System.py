# Author: --Wentao Ma--

'''
This project is designed for generating the ads document (.excel or .csv) that
could be recoginized by Google Ads System directly once you uploaded.
'''
from collections import OrderedDict
import xlsxwriter
import csv

# Campaign class defined all parameters
class Campaign:
    def __init__(self, name, Budget, Budget_type, type, Networks, languages, Bid_Strategy_Type, Start_Date, Ad_Schedule,
                 Ad_rotation, Delivery_method, Targeting_method, Exclusion_method, DSA_targeting_source):
        self.campaign_info = OrderedDict([('Campaign', ''), ('Budget','5'), ('Budget type','Daily'),('Campaign Type','Search'),
                                          ('Networks','Google search;Search Partners;Display Network'), ('Languages','en'),
                                          ('Bid Strategy Type','Maximize conversions'),('Start Date','2019-08-20'),
                                          ('Ad Schedule','(Monday[00:00-24:00]);(Tuesday[00:00-24:00]);(Wednesday[00:00-24:00]);(Thursday[00:00-24:00]);(Friday[00:00-24:00])'),
                                          ('Ad rotation','Optimize for clicks'),('Delivery method','Standard'),
                                          ('Targeting method','Location of presence or Area of interest'),
                                          ('Exclusion method','Location of presence or Area of interest'),
                                          ('DSA targeting source','Google'),('Ad Group',''),('Ad type',''),
                                          ('Headline 1',''),('Headline 2',''),('Headline 3',''),('Description Line 1',''),
                                          ('Description Line 2',''),('Final URL',''),('Max CPC',''),('Max CPM',''),
                                          ('Target CPM',''),('Keyword',''),('Criterion Type','')])
        self.campaign_info['Campaign'] = name
        self.campaign_info['Budget'] = Budget
        self.campaign_info['Budget type'] = Budget_type
        self.campaign_info['Campaign Type'] = type
        self.campaign_info['Networks'] = Networks
        self.campaign_info['Languages'] = languages
        self.campaign_info['Bid Strategy Type'] = Bid_Strategy_Type
        self.campaign_info['Start Date'] = Start_Date
        self.campaign_info['Ad Schedule'] = Ad_Schedule
        self.campaign_info['Ad rotation'] = Ad_rotation
        self.campaign_info['Delivery method'] = Delivery_method
        self.campaign_info['Targeting method'] = Targeting_method
        self.campaign_info['Exclusion method'] = Exclusion_method
        self.campaign_info['DSA targeting source'] = DSA_targeting_source
        # self.campaign_info['Ad Group'] = Ad_Group
        # self.campaign_info['Ad type'] = Ad_type
        # self.campaign_info['Headline 1'] = Headline1
        # self.campaign_info['Headline 2'] = Headline2
        # self.campaign_info['Headline 3'] = Headline3
        # self.campaign_info['Description Line 1'] = DescriptionLine1
        # self.campaign_info['Description Line 2'] = DescriptionLine2
        # self.campaign_info['Final URL'] = FinalURL
        # self.campaign_info['Max CPC'] = MaxCPC
        # self.campaign_info['Max CPM'] = MaxCPM
        # self.campaign_info['Target CPM'] = TargetCPM
        # self.campaign_info['Keyword'] = Keyword
        # self.campaign_info['Criterion Type'] = CriterionType

    def get_campaign_info(self):
        return self.campaign_info

# Ad Group class
class Adgroup:
    def __init__(self, name, Ad_Group, Ad_type, Headline1, Headline2, Headline3, DescriptionLine1, DescriptionLine2, FinalURL,
                 MaxCPC, MaxCPM, TargetCPM):
        self.adGroup_info = OrderedDict([('Campaign', ''), ('Budget',''), ('Budget type',''),('Campaign Type',''),
                                          ('Networks',''), ('Languages',''),('Bid Strategy Type',''),('Start Date',''),
                                          ('Ad Schedule',''),('Ad rotation',''),('Delivery method',''),('Targeting method',''),
                                          ('Exclusion method',''),('DSA targeting source',''),('Ad Group',''),('Ad type','Expanded text ad'),
                                          ('Headline 1',''),('Headline 2',''),('Headline 3','noviland.com | Simple Sourcing'),
                                          ('Description Line 1',''),('Description Line 2',''),('Final URL',''),('Max CPC','0.01'),
                                          ('Max CPM','0.01'), ('Target CPM','0.01'),('Keyword',''),('Criterion Type','')])

        self.adGroup_info['Campaign'] = name
        # self.adGroup_info['Budget'] = Budget
        # self.adGroup_info['Budget type'] = Budget_type
        # self.adGroup_info['Campaign Type'] = type
        # self.adGroup_info['Networks'] = Networks
        # self.adGroup_info['Languages'] = languages
        # self.adGroup_info['Bid Strategy Type'] = Bid_Strategy_Type
        # self.adGroup_info['Start Date'] = Start_Date
        # self.adGroup_info['Ad Schedule'] = Ad_Schedule
        # self.adGroup_info['Ad rotation'] = Ad_rotation
        # self.adGroup_info['Delivery method'] = Delivery_method
        # self.adGroup_info['Targeting method'] = Targeting_method
        # self.adGroup_info['Exclusion method'] = Exclusion_method
        # self.adGroup_info['DSA targeting source'] = DSA_targeting_source
        self.adGroup_info['Ad Group'] = Ad_Group
        self.adGroup_info['Ad type'] = Ad_type
        self.adGroup_info['Headline 1'] = Headline1
        self.adGroup_info['Headline 2'] = Headline2
        self.adGroup_info['Headline 3'] = Headline3
        self.adGroup_info['Description Line 1'] = DescriptionLine1
        self.adGroup_info['Description Line 2'] = DescriptionLine2
        self.adGroup_info['Final URL'] = FinalURL
        self.adGroup_info['Max CPC'] = MaxCPC
        self.adGroup_info['Max CPM'] = MaxCPM
        self.adGroup_info['Target CPM'] = TargetCPM
        # self.adGroup_info['Keyword'] = Keyword
        # self.adGroup_info['Criterion Type'] = CriterionType

    def get_adGroup_info(self):
        return self.adGroup_info

# keyword class
class keyword:
    def __init__(self,name, Ad_Group, Keyword, CriterionType):
        self.keyword_info = OrderedDict([('Campaign', ''), ('Budget',''), ('Budget type',''),('Campaign Type',''),
                                          ('Networks',''), ('Languages',''),('Bid Strategy Type',''),('Start Date',''),
                                          ('Ad Schedule',''),('Ad rotation',''),('Delivery method',''),('Targeting method',''),
                                          ('Exclusion method',''),('DSA targeting source',''),('Ad Group','Bakery Box'),('Ad type',''),
                                          ('Headline 1',''),('Headline 2',''),('Headline 3',''),('Description Line 1',''),
                                          ('Description Line 2',''),('Final URL',''),('Max CPC',''),('Max CPM',''),
                                          ('Target CPM',''),('Keyword',''),('Criterion Type','')])
        self.keyword_info['Campaign'] = name
        '''
        self.keyword_info['Budget'] = Budget
        self.keyword_info['Budget type'] = Budget_type
        self.keyword_info['Campaign Type'] = type
        self.keyword_info['Networks'] = Networks
        self.keyword_info['Languages'] = languages
        self.keyword_info['Bid Strategy Type'] = Bid_Strategy_Type
        self.keyword_info['Start Date'] = Start_Date
        self.keyword_info['Ad Schedule'] = Ad_Schedule
        self.keyword_info['Ad rotation'] = Ad_rotation
        self.keyword_info['Delivery method'] = Delivery_method
        self.keyword_info['Targeting method'] = Targeting_method
        self.keyword_info['Exclusion method'] = Exclusion_method
        self.keyword_info['DSA targeting source'] = DSA_targeting_source
        '''
        self.keyword_info['Ad Group'] = Ad_Group
        '''
        # self.keyword_info['Ad type'] = Ad_type
        # self.keyword_info['Headline 1'] = Headline1
        # self.keyword_info['Headline 2'] = Headline2
        # self.keyword_info['Headline 3'] = Headline3
        # self.keyword_info['Description Line 1'] = DescriptionLine1
        # self.keyword_info['Description Line 2'] = DescriptionLine2
        # self.keyword_info['Final URL'] = FinalURL
        # self.keyword_info['Max CPC'] = MaxCPC
        # self.keyword_info['Max CPM'] = MaxCPM
        # self.keyword_info['Target CPM'] = TargetCPM
        '''
        self.keyword_info['Keyword'] = Keyword
        self.keyword_info['Criterion Type'] = CriterionType

    def get_keyword_info(self):
        return self.keyword_info

if __name__ == '__main__':
    with open('/Users/wentaoma/PycharmProjects/untitled/Noviland/Noviland-Campaign-List') as read:
        reader = csv.reader(read)
        whole_List = [row for row in reader]
        # print(whole_List[1:])

    cur_campaign = 0
    C = Campaign(whole_List[1:][cur_campaign][0], 6, 'Daily', 'Search', 'Google search;Search Partners;Display Network',
                 'en', 'Maximize conversions', '2019-08-20',
                 '(Monday[00:00-24:00]);(Tuesday[00:00-24:00]);(Wednesday[00:00-24:00]);(Thursday[00:00-24:00]);(Friday[00:00-24:00',
                 'Optimize for clicks', 'Standard', 'Location of presence or Area of interest',
                 'Location of presence or Area of interest','Google')

    # 第一部分：当进入第一个Campaign的时候，先打印Campaign 成员变量类
    # First part: print the campaign line
    # Create four empty lists for storing the upcoming campaign information.
    Is_existed, campaign_results, result_1, campaign_res = [], [], [], []

    while cur_campaign < len(whole_List[1:]):
        if whole_List[1:][cur_campaign][0] not in Is_existed:
            Is_existed.append(whole_List[1:][cur_campaign][0])
            C = Campaign(whole_List[1:][cur_campaign][0], 6, 'Daily', 'Search',
                         'Google search;Search Partners;Display Network',
                         'en', 'Maximize conversions', '2019-08-20',
                         '(Monday[00:00-24:00]);(Tuesday[00:00-24:00]);(Wednesday[00:00-24:00]);(Thursday[00:00-24:00]);(Friday[00:00-24:00',
                         'Optimize for clicks', 'Standard', 'Location of presence or Area of interest',
                         'Location of presence or Area of interest',
                         'Google')
            res_1 = C.get_campaign_info()
            campaign_res.append(res_1)
        else:
            pass

    # Second part is used to print next four lines. (headline /description/ final url and something else)
    # 第二部分：紧接着打印 接下来4行的 Ad group 类，包括标题 两个描述 以及生成的网址链接
        Headline_2 = ['Simple Sourcing | Free Quotes','B2B Sourcing Factory Direct','B2B Sourcing Factory Direct','Easy B2B Purchase via Noviland']

        DescriptionLine_1 = ['Noviland makes buying ' + whole_List[1:][cur_campaign][3] +' from the most trusted factories easy',
                             'Noviland vets factories in China so you can easily source' + whole_List[1:][cur_campaign][3],
                             'Noviland only partners with the top ' + whole_List[1:][cur_campaign][2] + ' factories in China',
                             'Simplest way to purchase from ' + whole_List[1:][cur_campaign][2] +' factories | Risk-Free Pricing']

        DescriptionLine_2 = ['Factory selection based on your business and request. No credit card required. US-Support',
                             'Thousands of vetted factories.Zero - commitment.Factory direct pricing.B2B Purchasing',
                             'Risk-free sourcing from thousands of trusted factories. Zero commitment pricing. For SMBs',
                             'Risk-free sourcing from thousands of trusted factories. Zero commitment pricing. For SMBs']

        # 判断 Campaign name and Product name 中是否存在 空格
        # If there is a space between Campaign name and Product name.
        # 如果存在空格 我们需要将它进行替换成为'-' 来生成网址
        # If so, replace the space with '-' since we need to generate the final URL

        Campaign_name = whole_List[1:][cur_campaign][0]
        singular_product = whole_List[1:][cur_campaign][2]
        character = whole_List[1:][cur_campaign][1]

        if ' ' in singular_product:
            singular_product = singular_product.replace(' ', '-')
        else:
            pass

        if ' ' in Campaign_name:
            Campaign_name = Campaign_name.replace(' ', '-')
        else:
            pass

        Final_URL = 'https://noviland.com/alt-landing/?apd=sem-' + Campaign_name + '-' +singular_product + '-' + '01'

        line_1 = 0
        result_2 = []
        while line_1 < 4:
            if len(character) > 22:
                A = Adgroup(whole_List[1:][cur_campaign][0], whole_List[1:][cur_campaign][1], 'Expanded text ad',
                            whole_List[1:][cur_campaign][3],
                            Headline_2[line_1], 'noviland.com | Simple Sourcing', DescriptionLine_1[line_1],
                            DescriptionLine_2[line_1], Final_URL, 0.01, 0.01, 0.01)
            elif  20 < len(character) <= 22:
                A = Adgroup(whole_List[1:][cur_campaign][0], whole_List[1:][cur_campaign][1], 'Expanded text ad', whole_List[1:][cur_campaign][1] + ' Factory',
                        Headline_2[line_1] ,'noviland.com | Simple Sourcing',DescriptionLine_1[line_1], DescriptionLine_2[line_1], Final_URL, 0.01, 0.01, 0.01)
            else:
                A = Adgroup(whole_List[1:][cur_campaign][0], whole_List[1:][cur_campaign][1], 'Expanded text ad', whole_List[1:][cur_campaign][1] + ' Factories',
                        Headline_2[line_1] ,'noviland.com | Simple Sourcing',DescriptionLine_1[line_1], DescriptionLine_2[line_1], Final_URL, 0.01, 0.01, 0.01)
            res_2 = A.get_adGroup_info()
            campaign_res.append(res_2)
            line_1 += 1

        # 第三部分： 紧接着打印100多行的keywords类
        result_3 = []
        sep = '.'
        string = whole_List[1:][cur_campaign][1]
        string = string.replace(' ', '.')

        keyword_string = [whole_List[1:][cur_campaign][1] + ' b2b import',
                          whole_List[1:][cur_campaign][1] + ' b2b supply',
                          whole_List[1:][cur_campaign][1] + ' China Agent',
                          'China agent for ' + whole_List[1:][cur_campaign][1],
                          'Chinese agent for ' + whole_List[1:][cur_campaign][1],
                          whole_List[1:][cur_campaign][1] + ' China Sourcing Agent',
                          whole_List[1:][cur_campaign][1] + ' China Trade Agent',
                          'country that makes ' + whole_List[1:][cur_campaign][1],
                          'country that manufactures ' + whole_List[1:][cur_campaign][1],
                          'best country to buy ' + whole_List[1:][cur_campaign][1],
                          'Customs ' + whole_List[1:][cur_campaign][1],
                          'Customs broker ' + whole_List[1:][cur_campaign][1],
                          'Customs broker for ' + whole_List[1:][cur_campaign][1],
                          'Customs clearance ' + whole_List[1:][cur_campaign][1],
                          'Customs clearance FOR ' +whole_List[1:][cur_campaign][1],
                          'Directory for factory ' +whole_List[1:][cur_campaign][1],
                          'Directory for ' + whole_List[1:][cur_campaign][1] +' factories',
                          'Directory for ' + whole_List[1:][cur_campaign][1] +' suppliers',
                          'Directory for ' + whole_List[1:][cur_campaign][1] +' manufacturer',
                          'List of ' + whole_List[1:][cur_campaign][1] + ' factories',
                          'List of ' + whole_List[1:][cur_campaign][1] + ' suppliers',
                          'List of ' + whole_List[1:][cur_campaign][1] + ' manufacturers',
                          'Import duty ' + whole_List[1:][cur_campaign][1],
                          whole_List[1:][cur_campaign][1] +' import duties',
                          whole_List[1:][cur_campaign][1] +' factory',
                          whole_List[1:][cur_campaign][1] +' factory asia',
                          whole_List[1:][cur_campaign][1] +' factory China',
                          'factory that makes ' + whole_List[1:][cur_campaign][1],
                          whole_List[1:][cur_campaign][1] + ' factory vietnam',
                          whole_List[1:][cur_campaign][1] + ' Freight',
                          'good factory ' + whole_List[1:][cur_campaign][1],
                          'HS code ' + whole_List[1:][cur_campaign][1],
                          'HTS code' + whole_List[1:][cur_campaign][1],
                          whole_List[1:][cur_campaign][1] +' Importer',
                          whole_List[1:][cur_campaign][1] +' Importing',
                          whole_List[1:][cur_campaign][1] + ' Imports',
                          whole_List[1:][cur_campaign][1] +' Logistic Agent',
                          whole_List[1:][cur_campaign][1] +' manufacturer',
                          whole_List[1:][cur_campaign][1] +' manufacturer asia',
                          whole_List[1:][cur_campaign][1] +' manufacturer China',
                          whole_List[1:][cur_campaign][1] +' manufacturer vietnam',
                          'factory not in china ' +whole_List[1:][cur_campaign][1],
                          'outside china ' +whole_List[1:][cur_campaign][1],
                          'quality check ' +whole_List[1:][cur_campaign][1],
                          'quality control inspection for ' +whole_List[1:][cur_campaign][1],
                          'quality inspection ' + whole_List[1:][cur_campaign][1],
                          'Shipping ' + whole_List[1:][cur_campaign][1],
                          'Sourcing Agent ' +whole_List[1:][cur_campaign][1],
                          'Sourcing Agent for ' +whole_List[1:][cur_campaign][1],
                          whole_List[1:][cur_campaign][1]+ ' sourcing agent',
                          'Supplier China ' +whole_List[1:][cur_campaign][1],
                          'Tariff' + whole_List[1:][cur_campaign][1],
                          'Trade Agent' + whole_List[1:][cur_campaign][1],
                          'Trade Agent for ' + whole_List[1:][cur_campaign][1],
                          whole_List[1:][cur_campaign][1] +' trade agent',
                          'US Duty ' + whole_List[1:][cur_campaign][1],
                          'with the most factories for ' + whole_List[1:][cur_campaign][1],
                          'with the most manufacturers for ' + whole_List[1:][cur_campaign][1],
                          whole_List[1:][cur_campaign][1]+' manufacturer not in china',
                          whole_List[1:][cur_campaign][1]+ ' factory not in china',
                          '+b2b +import + ' +string,
                          '+b2b +supplier + '+string,
                          '+China +Agent + '+string,
                          '+China +Sourcing.Agent + ' +string,
                          '+China +Trade.Agent + '+string,
                          '+country.that.makes + '+'+'+string,
                          '+country.that.manufactures '+'+'+string,
                          '+Customs '+'+'+string,
                          '+Customs.broker '+'+'+string,
                          '+Customs.clearance '+'+'+string,
                          '+Directory.for +factory '+'+'+string,
                          '+US.Duty '+'+'+string,
                          '+factory '+'+'+string,
                          '+factory +asia '+'+'+string,
                          '+factory +China '+'+'+string,
                          '+factory.that.makes '+'+'+string,
                          '+factory +vietnam '+'+'+string,
                          '+Freight '+'+'+string,
                          '+Freight '+'+'+string,
                          '+good.factory '+'+'+string,
                          '+HS.code '+'+'+string,
                          '+HTS.code '+'+'+string,
                          '+Import '+'+'+string,
                          '+Importer '+'+'+string,
                          '+Importing '+'+'+string,
                          '+Imports '+'+'+string,
                          '+Logistic.Agent '+'+'+string,
                          '+manufacturer '+'+'+string,
                          '+manufacturer +asia '+'+'+string,
                          '+manufacturer +China '+'+'+string,
                          '+manufacturer +vietnam '+'+'+string,
                          '+factory.not.in.china '+'+'+string,
                          '+outside.china '+'+'+string,
                          '+outside.of.china '+'+'+string,
                          '+quality.check ' +'+'+ string,
                          '+quality.inspection '+'+'+string,
                          '+quality.control '+'+'+string,
                          '+Shipping '+'+'+string,
                          '+freight.forwarder '+'+'+string,
                          '+Sourcing.Agent '+'+'+string,
                          '+Supplier +China '+'+'+string,
                          '+Tariff '+'+'+string,
                          '+Trade.Agent '+'+'+string,
                          '+US Duty '+'+'+string,
                          '+with.the.most.manufacturers '+'+'+string,
                          '+manufacturer.not.in.china '+'+'+string,
                          '+factories.for '+'+'+string,
                          '+best.factories '+'+'+string,
                          '+wholesale.supplier '+'+'+string,
                          '+overseas.supplier '+'+'+string,
                          '+best.supplier '+'+'+string,
                          '+cheap.manufacturer '+'+'+string,
                          '+cheap +factory '+'+'+string,
                          '+factory.direct '+'+'+string,
                          '+big.manufacturer '+'+'+string,
                          '+largest.factory '+'+'+string,
                          '+biggest.supplier '+'+'+string,
                          '+biggest.factory '+'+'+string,
                          '+wholesaler '+'+'+string,
                          '+private.label '+'+'+string,
                          '+private.label '+'+'+string+' factory',
                          '+private.label '+'+'+string+' manufacturer',
                          '+private.label '+'+'+string +' supplier',
                          '+find '+string +'+'+' factories',
                          '+find '+string +'+'+' supplier',
                          '+find '+string +'+'+' manufacturers',
                          '+where.are '+'+'+string+' made']

        line_2 = 0
        while line_2 < 127:
            if line_2 < 60:
                k = keyword(whole_List[1:][cur_campaign][0], whole_List[1:][cur_campaign][2], keyword_string[line_2], 'Phrase')
                res_3 = k.get_keyword_info()
                campaign_res.append(res_3)

            else:
                k = keyword(whole_List[1:][cur_campaign][0], whole_List[1:][cur_campaign][2], keyword_string[line_2], 'Broad')
                res_3 = k.get_keyword_info()
                campaign_res.append(res_3)
            line_2 += 1
        cur_campaign += 1
    print(campaign_res)

# 第四部分: 生成新的CSV文件
# fourth part: this code is used to write the data into the new csv file
    with open('GoogleAds_Generator_version2.csv', 'w') as csvfile_final:
        fieldnames = ['Campaign', 'Budget', 'Budget type', 'Campaign Type', 'Networks', 'Languages','Bid Strategy Type',
                      'Start Date', 'Ad Schedule', 'Ad rotation', 'Delivery method', 'Targeting method','Exclusion method',
                      'DSA targeting source', 'Ad Group', 'Ad type', 'Headline 1', 'Headline 2', 'Headline 3', 'Description Line 1',
                      'Description Line 2', 'Final URL', 'Max CPC', 'Max CPM', 'Target CPM','Keyword', 'Criterion Type']

        writer = csv.DictWriter(csvfile_final, fieldnames=fieldnames)
        writer.writeheader()
        for row in campaign_res:
            writer.writerow(row)
