#-*-coding:utf-8-*-
import conf
from selenium.webdriver.support.ui import Select
from selenium.webdriver.common.alert import Alert 
from selenium import webdriver
import os,time,datetime
import xlrd
import xlwt
from xlutils.copy import copy

#初始化excel
def startExcel(bookName):
	book = xlwt.Workbook() 
	sheet = book.add_sheet(u"记录")
	sheet.write(0, 0, u"保健号")
	sheet.write(0, 1, u"姓名")
	sheet.write(0, 2, u"3个月体检")
	sheet.write(0, 3, u"6个月体检")
	sheet.write(0, 4, u"9个月体检")
	sheet.write(0, 5, u"12个月体检")
	sheet.write(0, 6, u"18个月体检")
	sheet.write(0, 7, u"24个月体检")
	sheet.write(0, 8, u"36个月体检")
	sheet.write(0, 9, u"说明：自动填写的记录将会有记录，未填写则没填写该记录。")
	book.save(bookName)
def edit(s_months,s_time):
	browser.find_element_by_name("chex_date").clear()

	if(s_months=='span_1'):
		#体检日期
		browser.find_element_by_name("chex_date").send_keys(s_time)
		#体重
		browser.find_element_by_name("chex_weigth").clear()
		browser.find_element_by_name("chex_weigth").send_keys("6.4")
		#身高
		browser.find_element_by_name("chex_height").clear()
		browser.find_element_by_name("chex_height").send_keys("61.5")
		#胸围
		browser.find_element_by_name("chex_bust").clear()
		browser.find_element_by_name("chex_bust").send_keys("41")
		#头围
		browser.find_element_by_name("chex_head_circumference").clear()
		browser.find_element_by_name("chex_head_circumference").send_keys("42")
		#前囟门
		browser.find_element_by_name("chex_before_fontanel1").clear()
		browser.find_element_by_name("chex_before_fontanel1").send_keys("2")
		browser.find_element_by_name("chex_before_fontanel2").clear()
		browser.find_element_by_name("chex_before_fontanel2").send_keys("2")	
		#下拉按钮
		browser.find_element_by_name("chex_feeding_situation").click()
		browser.find_element_by_xpath("//option[@value='1688']").click()

		browser.find_element_by_name("chex_feeding_practices").click()
		browser.find_element_by_xpath("//option[@value='1690']").click()

		browser.find_element_by_name("chex_complementary").click()
		browser.find_element_by_xpath("//option[@value='1717']").click()
		#牙齿
		browser.find_element_by_name("chex_teeth").clear()
		browser.find_element_by_name("chex_teeth").send_keys("0")

		#下拉
		select = Select(browser.find_element_by_name("chex_ric")) 
		select.select_by_visible_text(u"无")

		select = Select(browser.find_element_by_name("chex_signs")) 
		select.select_by_visible_text(u"无")	
		
		select = Select(browser.find_element_by_name("chex_history")) 
		select.select_by_visible_text(u"无")		


		#指导事项其他
		browser.find_element_by_name("chex_guidance_other").clear()
		browser.find_element_by_name("chex_guidance_other").send_keys(u"查验血常规、骨碱性磷酸酶，配方奶每天3-4次（约800毫升），吃厚粥，碎菜，面食等；训练扶站、迈步、站立，语言；口腔清洁；多晒太阳，补充维生素D，必要时补充钙剂；按时预防接种，建议添加营养包。")		
		#点击评价
		browser.find_element_by_name("conclude_button").click()
		#保存
		browser.find_element_by_xpath("/html/body/form/table/tbody/tr[3]/td/table/tbody/tr[2]/td/input[1]").click()
		time.sleep(3)
		Alert(browser).accept()
		#点击添加
		browser.find_element_by_name("add").click()
	if(s_months=='span_2'):
		#体检日期
		browser.find_element_by_name("chex_date").send_keys(s_time)
		# #孤独症筛查检验
		# browser.find_element_by_name("chex_autism_div_ny").click()
		# browser.find_element_by_xpath("//option[@value='2']").click()
		#体重
		browser.find_element_by_name("chex_weigth").clear()
		browser.find_element_by_name("chex_weigth").send_keys("7.9")
		#身高
		browser.find_element_by_name("chex_height").clear()
		browser.find_element_by_name("chex_height").send_keys("68")
		#胸围
		browser.find_element_by_name("chex_bust").clear()
		browser.find_element_by_name("chex_bust").send_keys("42")
		#头围
		browser.find_element_by_name("chex_head_circumference").clear()
		browser.find_element_by_name("chex_head_circumference").send_keys("43")
		#前囟门
		browser.find_element_by_name("chex_before_fontanel1").clear()
		browser.find_element_by_name("chex_before_fontanel1").send_keys("1.5")
		browser.find_element_by_name("chex_before_fontanel2").clear()
		browser.find_element_by_name("chex_before_fontanel2").send_keys("1.5")	
		#下拉按钮
		#喂养情况
		browser.find_element_by_name("chex_feeding_situation").click()
		browser.find_element_by_xpath("//option[@value='1688']").click()
		#喂养方法
		browser.find_element_by_name("chex_feeding_practices").click()
		browser.find_element_by_xpath("//option[@value='1691']").click()
		#添加辅食
		browser.find_element_by_name("chex_complementary").click()
		browser.find_element_by_xpath("//option[@value='1716']").click()
		#牙齿
		browser.find_element_by_name("chex_teeth").clear()
		browser.find_element_by_name("chex_teeth").send_keys("2")
		#听力筛查
		browser.find_element_by_name("chex_hearing_screening").click()
		browser.find_element_by_xpath("//option[@value='350']").click()
		#眼睛
		browser.find_element_by_name("chex_eye").click()
		browser.find_element_by_xpath("//option[@value='1718']").click()
		#视力筛查
		select = Select(browser.find_element_by_name("chex_vision_screening")) 
		select.select_by_visible_text(u"正常")
		#血红蛋白检查
		browser.find_element_by_name("chex_hemoglobin_checks").send_keys("121")
		#佝偻症病状
		select = Select(browser.find_element_by_name("chex_ric")) 
		select.select_by_visible_text(u"无")
		#佝偻症体征
		select = Select(browser.find_element_by_name("chex_signs")) 
		select.select_by_visible_text(u"无")	
		#佝偻症病史
		select = Select(browser.find_element_by_name("chex_history")) 
		select.select_by_visible_text(u"无")		
		#家里同胞患有孤独症
		select = Select(browser.find_element_by_name("chex_asd_diagnoses")) 
		select.select_by_visible_text(u"无")	
		#担心孩子发育问题
		select = Select(browser.find_element_by_name("chex_worry_son")) 
		select.select_by_visible_text(u"无")
		#12-24月孩子发育倒退	
		select = Select(browser.find_element_by_name("chex_hypogenesis12")) 
		select.select_by_visible_text(u"无")
		#18至24月龄是否做过筛查 
		select = Select(browser.find_element_by_name("chex_is_screening18")) 
		select.select_by_visible_text(u"有")
		#孤独者一级检查结论
		select = Select(browser.find_element_by_name("chex_autism_conclusion")) 
		select.select_by_visible_text(u"正常")

		#指导事项其他
		browser.find_element_by_name("chex_guidance_other").clear()
		browser.find_element_by_name("chex_guidance_other").send_keys(u"查验血常规、骨碱性磷酸酶。坚持母乳喂养，逐步添加强化铁辅食等，如果汁、米粉、烂粥、鱼、肝泥、肉末等；训练靠坐、独坐；多晒太阳；补充维生素D，必要时补充钙剂；按时预防接种。建议添加营养包")		
		#点击评价
		browser.find_element_by_name("conclude_button").click()
		#弹出框
		browser.find_element_by_xpath("/html/body/form/table/tbody/tr[3]/td/table/tbody/tr[2]/td/input[1]").click()
		time.sleep(3)
		Alert(browser).dismiss()
		time.sleep(3)
		Alert(browser).accept()
		#点击添加
		browser.find_element_by_name("add").click()
	if(s_months=='span_3'):
		#体检日期
		browser.find_element_by_name("chex_date").send_keys(s_time)
		# #孤独症筛查检验
		# browser.find_element_by_name("chex_autism_div_ny").click()
		# browser.find_element_by_xpath("//option[@value='2']").click()
		#体重
		browser.find_element_by_name("chex_weigth").clear()
		browser.find_element_by_name("chex_weigth").send_keys("8.9")
		#身高
		browser.find_element_by_name("chex_height").clear()
		browser.find_element_by_name("chex_height").send_keys("72")
		#胸围
		browser.find_element_by_name("chex_bust").clear()
		browser.find_element_by_name("chex_bust").send_keys("43.5")
		#头围
		browser.find_element_by_name("chex_head_circumference").clear()
		browser.find_element_by_name("chex_head_circumference").send_keys("44")
		#前囟门
		browser.find_element_by_name("chex_before_fontanel1").clear()
		browser.find_element_by_name("chex_before_fontanel1").send_keys("0.5")
		browser.find_element_by_name("chex_before_fontanel2").clear()
		browser.find_element_by_name("chex_before_fontanel2").send_keys("0.5")	
		#下拉按钮
		#喂养情况
		browser.find_element_by_name("chex_feeding_situation").click()
		browser.find_element_by_xpath("//option[@value='1688']").click()
		#牙齿
		browser.find_element_by_name("chex_teeth").clear()
		browser.find_element_by_name("chex_teeth").send_keys("4")
		#听力筛查
		browser.find_element_by_name("chex_hearing_screening").click()
		browser.find_element_by_xpath("//option[@value='350']").click()
		#眼睛
		browser.find_element_by_name("chex_eye").click()
		browser.find_element_by_xpath("//option[@value='1718']").click()
		#视力筛查
		select = Select(browser.find_element_by_name("chex_vision_screening")) 
		select.select_by_visible_text(u"正常")
		#血红蛋白检查
		browser.find_element_by_name("chex_hemoglobin_checks").send_keys("124")
		#智商
		select = Select(browser.find_element_by_name("chex_iq")) 
		select.select_by_visible_text(u"正常")


		#佝偻症病状
		select = Select(browser.find_element_by_name("chex_ric")) 
		select.select_by_visible_text(u"无")
		#佝偻症体征
		select = Select(browser.find_element_by_name("chex_signs")) 
		select.select_by_visible_text(u"无")	
		#佝偻症病史
		select = Select(browser.find_element_by_name("chex_history")) 
		select.select_by_visible_text(u"无")		
		#家里同胞患有孤独症
		select = Select(browser.find_element_by_name("chex_asd_diagnoses")) 
		select.select_by_visible_text(u"无")	
		#担心孩子发育问题
		select = Select(browser.find_element_by_name("chex_worry_son")) 
		select.select_by_visible_text(u"无")
		#孤独者一级检查结论
		select = Select(browser.find_element_by_name("chex_autism_conclusion")) 
		select.select_by_visible_text(u"正常")
		#指导事项其他
		browser.find_element_by_name("chex_guidance_other").clear()
		browser.find_element_by_name("chex_guidance_other").send_keys(u"坚持母乳喂养，不足添加配方奶，减少喂奶的次数，添加米粉、肉末、蛋羹、碎菜、烂面等；练习爬行站立；注意安全；多晒太阳，补充维生素D，必要时补充钙剂；按时预防接种，建议添加营养包")		
		#点击评价
		browser.find_element_by_name("conclude_button").click()
		browser.find_element_by_xpath("/html/body/form/table/tbody/tr[3]/td/table/tbody/tr[2]/td/input[1]").click()
		time.sleep(3)
		Alert(browser).dismiss()
		time.sleep(3)
		Alert(browser).accept()
		#点击添加
		browser.find_element_by_name("add").click()
	if(s_months=='span_4'):
		#体检日期
		browser.find_element_by_name("chex_date").send_keys(s_time)
		# #孤独症筛查检验
		# browser.find_element_by_name("chex_autism_div_ny").click()
		# browser.find_element_by_xpath("//option[@value='2']").click()
		#体重
		browser.find_element_by_name("chex_weigth").clear()
		browser.find_element_by_name("chex_weigth").send_keys("9.6")
		#身高
		browser.find_element_by_name("chex_height").clear()
		browser.find_element_by_name("chex_height").send_keys("76")
		#胸围
		browser.find_element_by_name("chex_bust").clear()
		browser.find_element_by_name("chex_bust").send_keys("44.5")
		#头围
		browser.find_element_by_name("chex_head_circumference").clear()
		browser.find_element_by_name("chex_head_circumference").send_keys("45")
		#前囟门
		browser.find_element_by_name("chex_before_fontanel1").clear()
		browser.find_element_by_name("chex_before_fontanel1").send_keys("0")
		browser.find_element_by_name("chex_before_fontanel2").clear()
		browser.find_element_by_name("chex_before_fontanel2").send_keys("0")	
		#下拉按钮
		#喂养情况
		browser.find_element_by_name("chex_feeding_situation").click()
		browser.find_element_by_xpath("//option[@value='1688']").click()
		#断奶
		select = Select(browser.find_element_by_name("chex_weaning")) 
		select.select_by_visible_text(u"是")
		#牙齿
		browser.find_element_by_name("chex_teeth").clear()
		browser.find_element_by_name("chex_teeth").send_keys("6")
		#龋齿
		select = Select(browser.find_element_by_name("chex_caries")) 
		select.select_by_visible_text(u"无")
		#牙齿清洁
		select = Select(browser.find_element_by_name("chex_clean_teeth")) 
		select.select_by_visible_text(u"清洁")
		#咬合畸形
		select = Select(browser.find_element_by_name("chex_abnormal_occlusion")) 
		select.select_by_visible_text(u"无")
		#听力筛查
		browser.find_element_by_name("chex_hearing_screening").click()
		browser.find_element_by_xpath("//option[@value='350']").click()
		#眼睛
		browser.find_element_by_name("chex_eye").click()
		browser.find_element_by_xpath("//option[@value='1718']").click()
		#视力筛查
		select = Select(browser.find_element_by_name("chex_vision_screening")) 
		select.select_by_visible_text(u"正常")
		#血红蛋白检查
		browser.find_element_by_name("chex_hemoglobin_checks").send_keys("122")
		#智商
		# select = Select(browser.find_element_by_name("chex_iq")) 
		# select.select_by_visible_text(u"正常")


		#佝偻症病状
		select = Select(browser.find_element_by_name("chex_ric")) 
		select.select_by_visible_text(u"无")
		#佝偻症体征
		select = Select(browser.find_element_by_name("chex_signs")) 
		select.select_by_visible_text(u"无")	
		#佝偻症病史
		select = Select(browser.find_element_by_name("chex_history")) 
		select.select_by_visible_text(u"无")		
		#家里同胞患有孤独症
		select = Select(browser.find_element_by_name("chex_asd_diagnoses")) 
		select.select_by_visible_text(u"无")	
		#担心孩子发育问题
		select = Select(browser.find_element_by_name("chex_worry_son")) 
		select.select_by_visible_text(u"无")
		#12-24月孩子发育倒退	
		select = Select(browser.find_element_by_name("chex_hypogenesis12")) 
		select.select_by_visible_text(u"无")
		#18至24月龄是否做过筛查 
		# select = Select(browser.find_element_by_name("chex_is_screening18")) 
		# select.select_by_visible_text(u"有")
		#孤独者一级检查结论
		select = Select(browser.find_element_by_name("chex_autism_conclusion")) 
		select.select_by_visible_text(u"正常")
		#指导事项其他
		browser.find_element_by_name("chex_guidance_other").clear()
		browser.find_element_by_name("chex_guidance_other").send_keys(u"查验血常规、骨碱性磷酸酶，配方奶每天3-4次（约800毫升），吃厚粥，碎菜，面食等；训练扶站、迈步、站立，语言；口腔清洁；多晒太阳，补充维生素D，必要时补充钙剂；按时预防接种，建议添加营养包。")		
		#点击评价
		browser.find_element_by_name("conclude_button").click()
		#保存
		browser.find_element_by_xpath("/html/body/form/table/tbody/tr[3]/td/table/tbody/tr[2]/td/input[1]").click()
		time.sleep(3)
		Alert(browser).dismiss()
		time.sleep(3)
		Alert(browser).accept()
		#点击添加
		browser.find_element_by_name("add").click()
	if(s_months=='span_5'):
		#体检日期
		browser.find_element_by_name("chex_date").send_keys(s_time)
		# #孤独症筛查检验
		# browser.find_element_by_name("chex_autism_div_ny").click()
		# browser.find_element_by_xpath("//option[@value='2']").click()
		#体重
		browser.find_element_by_name("chex_weigth").clear()
		browser.find_element_by_name("chex_weigth").send_keys("11")
		#身高
		browser.find_element_by_name("chex_height").clear()
		browser.find_element_by_name("chex_height").send_keys("83")
		#胸围
		browser.find_element_by_name("chex_bust").clear()
		browser.find_element_by_name("chex_bust").send_keys("47")
		#头围
		browser.find_element_by_name("chex_head_circumference").clear()
		browser.find_element_by_name("chex_head_circumference").send_keys("47")
		#前囟门
		browser.find_element_by_name("chex_before_fontanel1").clear()
		browser.find_element_by_name("chex_before_fontanel1").send_keys("0")
		browser.find_element_by_name("chex_before_fontanel2").clear()
		browser.find_element_by_name("chex_before_fontanel2").send_keys("0")
		#囟门
		select = Select(browser.find_element_by_name("chex_fontanel")) 
		select.select_by_visible_text(u"已闭")
		#喂养情况
		browser.find_element_by_name("chex_feeding_situation").click()
		browser.find_element_by_xpath("//option[@value='1688']").click()
		#断奶
		select = Select(browser.find_element_by_name("chex_weaning")) 
		select.select_by_visible_text(u"是")
		#牙齿
		browser.find_element_by_name("chex_teeth").clear()
		browser.find_element_by_name("chex_teeth").send_keys("14")
		#龋齿
		select = Select(browser.find_element_by_name("chex_caries")) 
		select.select_by_visible_text(u"无")
		#牙齿清洁
		select = Select(browser.find_element_by_name("chex_clean_teeth")) 
		select.select_by_visible_text(u"清洁")
		#咬合畸形
		select = Select(browser.find_element_by_name("chex_abnormal_occlusion")) 
		select.select_by_visible_text(u"无")
		#听力筛查
		browser.find_element_by_name("chex_hearing_screening").click()
		browser.find_element_by_xpath("//option[@value='350']").click()
		#眼睛
		browser.find_element_by_name("chex_eye").click()
		browser.find_element_by_xpath("//option[@value='1718']").click()
		#视力筛查
		select = Select(browser.find_element_by_name("chex_vision_screening")) 
		select.select_by_visible_text(u"正常")
		#血红蛋白检查
		browser.find_element_by_name("chex_hemoglobin_checks").send_keys("125")
		#智商
		select = Select(browser.find_element_by_name("chex_iq")) 
		select.select_by_visible_text(u"正常")


		#佝偻症病状
		select = Select(browser.find_element_by_name("chex_ric")) 
		select.select_by_visible_text(u"无")
		#佝偻症体征
		select = Select(browser.find_element_by_name("chex_signs")) 
		select.select_by_visible_text(u"无")	
		#佝偻症病史
		select = Select(browser.find_element_by_name("chex_history")) 
		select.select_by_visible_text(u"无")		
		#家里同胞患有孤独症
		select = Select(browser.find_element_by_name("chex_asd_diagnoses")) 
		select.select_by_visible_text(u"无")	
		#担心孩子发育问题
		select = Select(browser.find_element_by_name("chex_worry_son")) 
		select.select_by_visible_text(u"无")
		#12-24月孩子发育倒退	
		select = Select(browser.find_element_by_name("chex_hypogenesis12")) 
		select.select_by_visible_text(u"无")
		#18至24月龄是否做过筛查 
		select = Select(browser.find_element_by_name("chex_is_screening18")) 
		select.select_by_visible_text(u"有")
		#孤独者一级检查结论
		select = Select(browser.find_element_by_name("chex_autism_conclusion")) 
		select.select_by_visible_text(u"正常")
		#指导事项其他
		browser.find_element_by_name("chex_guidance_other").clear()
		browser.find_element_by_name("chex_guidance_other").send_keys(u"适量配方奶，软饭碎菜，注意喂养和饮食习惯，均衡饮食；防止意外伤害；口腔清洁；关注儿童心理健康；加强户外活动，社交、语言训练；多晒太阳，补充维生素D，必要时补充钙剂；按时预防接种，建议添加营养包")		
		#点击评价
		browser.find_element_by_name("conclude_button").click()
		#保存
		browser.find_element_by_xpath("/html/body/form/table/tbody/tr[3]/td/table/tbody/tr[2]/td/input[1]").click()
		time.sleep(3)
		Alert(browser).dismiss()
		time.sleep(3)
		Alert(browser).accept()
		#点击添加
		browser.find_element_by_name("add").click()
	if(s_months=='span_6'):
		#体检日期
		browser.find_element_by_name("chex_date").send_keys(s_time)
		# #孤独症筛查检验
		# browser.find_element_by_name("chex_autism_div_ny").click()
		# browser.find_element_by_xpath("//option[@value='2']").click()
		#体重
		browser.find_element_by_name("chex_weigth").clear()
		browser.find_element_by_name("chex_weigth").send_keys("12.8")
		#身高
		browser.find_element_by_name("chex_height").clear()
		browser.find_element_by_name("chex_height").send_keys("88")
		#胸围
		browser.find_element_by_name("chex_bust").clear()
		browser.find_element_by_name("chex_bust").send_keys("50")
		#头围
		browser.find_element_by_name("chex_head_circumference").clear()
		browser.find_element_by_name("chex_head_circumference").send_keys("48")
		#前囟门
		# browser.find_element_by_name("chex_before_fontanel1").clear()
		# browser.find_element_by_name("chex_before_fontanel1").send_keys("0")
		# browser.find_element_by_name("chex_before_fontanel2").clear()
		# browser.find_element_by_name("chex_before_fontanel2").send_keys("0")
		#囟门
		select = Select(browser.find_element_by_name("chex_fontanel")) 
		select.select_by_visible_text(u"已闭")
		#喂养情况
		# browser.find_element_by_name("chex_feeding_situation").click()
		# browser.find_element_by_xpath("//option[@value='1688']").click()
		#断奶
		# select = Select(browser.find_element_by_name("chex_weaning")) 
		# select.select_by_visible_text(u"是")
		#牙齿
		browser.find_element_by_name("chex_teeth").clear()
		browser.find_element_by_name("chex_teeth").send_keys("16")
		#龋齿
		select = Select(browser.find_element_by_name("chex_caries")) 
		select.select_by_visible_text(u"无")
		#牙齿清洁
		select = Select(browser.find_element_by_name("chex_clean_teeth")) 
		select.select_by_visible_text(u"清洁")
		#咬合畸形
		select = Select(browser.find_element_by_name("chex_abnormal_occlusion")) 
		select.select_by_visible_text(u"无")
		#听力筛查
		browser.find_element_by_name("chex_hearing_screening").click()
		browser.find_element_by_xpath("//option[@value='350']").click()
		#眼睛
		browser.find_element_by_name("chex_eye").click()
		browser.find_element_by_xpath("//option[@value='1718']").click()
		#视力筛查
		select = Select(browser.find_element_by_name("chex_vision_screening")) 
		select.select_by_visible_text(u"正常")
		#血红蛋白检查
		browser.find_element_by_name("chex_hemoglobin_checks").send_keys("128")
		#智商
		# select = Select(browser.find_element_by_name("chex_iq")) 
		# select.select_by_visible_text(u"正常")


		#佝偻症病状
		select = Select(browser.find_element_by_name("chex_ric")) 
		select.select_by_visible_text(u"无")
		#佝偻症体征
		select = Select(browser.find_element_by_name("chex_signs")) 
		select.select_by_visible_text(u"无")	
		#佝偻症病史
		select = Select(browser.find_element_by_name("chex_history")) 
		select.select_by_visible_text(u"无")		
		#家里同胞患有孤独症
		select = Select(browser.find_element_by_name("chex_asd_diagnoses")) 
		select.select_by_visible_text(u"无")	
		#担心孩子发育问题
		select = Select(browser.find_element_by_name("chex_worry_son")) 
		select.select_by_visible_text(u"无")
		#12-24月孩子发育倒退	
		select = Select(browser.find_element_by_name("chex_hypogenesis12")) 
		select.select_by_visible_text(u"无")
		#18至24月龄是否做过筛查 
		select = Select(browser.find_element_by_name("chex_is_screening18")) 
		select.select_by_visible_text(u"有")
		#孤独者一级检查结论
		select = Select(browser.find_element_by_name("chex_autism_conclusion")) 
		select.select_by_visible_text(u"正常")
		#指导事项其他
		browser.find_element_by_name("chex_guidance_other").clear()
		browser.find_element_by_name("chex_guidance_other").send_keys(u"查验血常规、骨碱性磷酸酶。防止孩子偏食挑食，建立均衡饮食概念；运动语言训练安全教育，防止意外伤害；口腔卫生；培养独立性，关注儿童心理健康，对孩子不要溺爱；按时预防接种，建议添加营养包。")		
		#点击评价
		browser.find_element_by_name("conclude_button").click()
		#保存
		browser.find_element_by_xpath("/html/body/form/table/tbody/tr[3]/td/table/tbody/tr[2]/td/input[1]").click()	
		time.sleep(3)
		Alert(browser).dismiss()
		time.sleep(3)
		Alert(browser).accept()
		#点击添加
		browser.find_element_by_name("add").click()
	if(s_months=='span_7'):
		#体检日期
		browser.find_element_by_name("chex_date").send_keys(s_time)
		# #孤独症筛查检验
		# browser.find_element_by_name("chex_autism_div_ny").click()
		# browser.find_element_by_xpath("//option[@value='2']").click()
		#体重
		browser.find_element_by_name("chex_weigth").clear()
		browser.find_element_by_name("chex_weigth").send_keys("14.8")
		#身高
		browser.find_element_by_name("chex_height").clear()
		browser.find_element_by_name("chex_height").send_keys("97")
		#胸围
		browser.find_element_by_name("chex_bust").clear()
		browser.find_element_by_name("chex_bust").send_keys("51.5")
		#头围
		browser.find_element_by_name("chex_head_circumference").clear()
		browser.find_element_by_name("chex_head_circumference").send_keys("50")
		#龋齿
		select = Select(browser.find_element_by_name("chex_caries")) 
		select.select_by_visible_text(u"无")
		#牙齿清洁
		select = Select(browser.find_element_by_name("chex_clean_teeth")) 
		select.select_by_visible_text(u"清洁")
		#咬合畸形
		select = Select(browser.find_element_by_name("chex_abnormal_occlusion")) 
		select.select_by_visible_text(u"无")
		#听力筛查
		browser.find_element_by_name("chex_hearing_screening").click()
		browser.find_element_by_xpath("//option[@value='350']").click()
		#眼睛
		browser.find_element_by_name("chex_eye").click()
		browser.find_element_by_xpath("//option[@value='1718']").click()
		#视力筛查
		select = Select(browser.find_element_by_name("chex_vision_screening")) 
		select.select_by_visible_text(u"正常")
		#血红蛋白检查
		browser.find_element_by_name("chex_hemoglobin_checks").send_keys("126")
		#视力左边
		browser.find_element_by_name("chex_left_vision").clear()
		browser.find_element_by_name("chex_left_vision").send_keys("5.1")
		#视力右边
		browser.find_element_by_name("chex_right_vision").clear()
		browser.find_element_by_name("chex_right_vision").send_keys("5.1")
		#指导事项其他
		browser.find_element_by_name("chex_guidance_other").clear()
		browser.find_element_by_name("chex_guidance_other").send_keys(u"防止孩子偏食挑食，建立均衡饮食概念；运动语言训练，鼓励孩子与同龄孩子和家人交流；防止意外伤害；口腔卫生；培养独立性，关注儿童心理健康，对孩子不要溺爱；按时预防接种，建议添加营养包。")		
		#点击评价
		browser.find_element_by_name("conclude_button").click()
		#保存
		browser.find_element_by_xpath("/html/body/form/table/tbody/tr[3]/td/table/tbody/tr[2]/td/input[1]").click()
		time.sleep(3)
		Alert(browser).accept()
		browser.implicitly_wait(3)
		#点击返回
		browser.find_element_by_name("return").click()
	return None
def check(m,excelName):
	m = (m-1)*10+1
	#选择人
	for i in ["tr[2]","tr[3]","tr[4]","tr[5]","tr[6]","tr[7]","tr[8]","tr[9]","tr[10]","tr[11]"]:
		#保健号
		x_pathNum = "/html/body/form/table/tbody/tr[6]/td/table/tbody/"+i+"/td[1]"
		number = browser.find_element_by_xpath(x_pathNum).text
		#姓名
		x_pathName = "/html/body/form/table/tbody/tr[6]/td/table/tbody/"+i+"/td[2]"
		name = browser.find_element_by_xpath(x_pathName).text
		#保存保健号，姓名到excel
		oldWb = xlrd.open_workbook(excelName, formatting_info=True);
		newWb = copy(oldWb)
		newWs = newWb.get_sheet(0)
		newWs.write(m, 0, number)
		newWs.write(m, 1, name)
		newWb.save(excelName)
		# excel_Name = time.strftime('%Y-%m-%d-%H-%M-%S',time.localtime(time.time()))
		# newWb.save(excel_Name)
		#点击 增加按钮
		x_path3 = "/html/body/form/table/tbody/tr[6]/td/table/tbody/"+i+"/td[9]/a[1]"
		browser.find_element_by_xpath(x_path3).click()
		#browser.implicitly_wait(2)
		#browser.find_element_by_xpath("/html/body/form/table/tbody/tr[6]/td/table/tbody/tr[3]/td[9]/a[1]").click()
		#                                /html/body/form/table/tbody/tr[6]/td/table/tbody/tr[7]/td[9]/a[1]
			#s_check为未体检 全局变量
		try:
			browser.switch_to_alert().accept()
		except:
			s_check = u"未体检"

			#判断当前日期之前是否完成体检 s_date2为当前日期
			for num in ['span_1','span_2','span_3','span_4','span_5','span_6','span_7']:
				s_date = browser.find_element_by_xpath("//*[@id='span_1']/table/tbody/tr/td[1]").text
				#x_path1 为日期  now 为当前日期
				x_path1 = "//*[@id='"+num+"']/table/tbody/tr/td[1]"
				s_date = browser.find_element_by_xpath(x_path1).text
				#now = datetime.datetime.now().strftime("%Y-%m-%d")
				s_date1 = datetime.datetime.strptime(s_date,'%Y-%m-%d')
				s_date2 = time.strftime('%Y-%m-%d',time.localtime(time.time()))
				s_date2 = datetime.datetime.strptime(s_date2,'%Y-%m-%d')
				delta = s_date2 - s_date1
				#print delta.days
				x_path2 = "//*[@id='"+num+"']/table/tbody/tr/td[5]"
				check = browser.find_element_by_xpath(x_path2).text
				# print "#############################"
				# print u"正在检查日期：",s_date1
				if(delta.days >= 0 ):
					if(check == s_check):
						edit(num,s_date)
						print "----------ok----------",num
						print u"添加"+s_date+u"体检记录"
						#添加到excel记录
						addExcel(m,num,excelName)						
					elif (num == "span_7"):
						browser.find_element_by_name("button").click()
						Alert(browser).accept()
			time.sleep(3)
			print i+u"检查完毕"
			m = m+1
def addExcel(m,num,excelName):
	if (num == "span_1"):
		#保存检查记录
		oldWb = xlrd.open_workbook(excelName, formatting_info=True);
		newWb = copy(oldWb)
		newWs = newWb.get_sheet(0)
		newWs.write(m, 2, u"添加体检记录")
		newWb.save(excelName)
	if (num == "span_2"):
		#保存检查记录
		oldWb = xlrd.open_workbook(excelName, formatting_info=True);
		newWb = copy(oldWb)
		newWs = newWb.get_sheet(0)
		newWs.write(m, 3, u"添加体检记录")
		newWb.save(excelName)		
	if (num == "span_3"):
		#保存检查记录
		oldWb = xlrd.open_workbook(excelName, formatting_info=True);
		newWb = copy(oldWb)
		newWs = newWb.get_sheet(0)
		newWs.write(m, 4, u"添加体检记录")
		newWb.save(excelName)
	if (num == "span_4"):
		#保存检查记录
		oldWb = xlrd.open_workbook(excelName, formatting_info=True);
		newWb = copy(oldWb)
		newWs = newWb.get_sheet(0)
		newWs.write(m, 5, u"添加体检记录")
		newWb.save(excelName)
	if (num == "span_5"):
		#保存检查记录
		oldWb = xlrd.open_workbook(excelName, formatting_info=True);
		newWb = copy(oldWb)
		newWs = newWb.get_sheet(0)
		newWs.write(m, 6, u"添加体检记录")
		newWb.save(excelName)
	if (num == "span_6"):
		#保存检查记录
		oldWb = xlrd.open_workbook(excelName, formatting_info=True);
		newWb = copy(oldWb)
		newWs = newWb.get_sheet(0)
		newWs.write(m, 7, u"添加体检记录")
		newWb.save(excelName)
	if (num == "span_7"):
		#保存检查记录
		oldWb = xlrd.open_workbook(excelName, formatting_info=True);
		newWb = copy(oldWb)
		newWs = newWb.get_sheet(0)
		newWs.write(m, 8, u"添加体检记录")
		newWb.save(excelName)

#登录
browser = webdriver.Ie()
browser.maximize_window()     #最大化
browser.get("http://218.18.233.230:8888")

#读取配置文件conf.xls
book = xlrd.open_workbook("conf.xls")
sh = book.sheet_by_index(0)
userId = sh.cell_value(rowx=0,colx=1)
passWord = sh.cell_value(rowx=1,colx=1)
timeStart = sh.cell_value(rowx=3,colx=1)
timeEnd = sh.cell_value(rowx=3,colx=2)

browser.find_element_by_name("userId").send_keys(userId)
browser.find_element_by_name("password").send_keys(passWord)
time.sleep(6)
browser.find_element_by_name("submit").click()


#提示框
browser.implicitly_wait(5)
#Alert(browser).accept()

#跳转页面
# browser.implicitly_wait(1.5)
browser.get("http://218.18.233.230:8888/index3.do?system=ChildInfo")
browser.implicitly_wait(10)

#选择菜单
time.sleep(2)
browser.find_element_by_id("mMenu0").click()
# browser.implicitly_wait(10)
time.sleep(2)
browser.find_element_by_xpath("//*[@id='mmenudiv0']/table/tbody/tr[2]/td").click()
# browser.implicitly_wait(10)
# time.sleep(2)
browser.switch_to_frame("mainWorkArea")

#选择 保健号 来查询   0001010220000067389
# browser.find_element_by_name("healthno").clear()
# browser.find_element_by_name("healthno").send_keys("0001010220000036806")
#选择 出生时间段 来查询
browser.find_element_by_name("birth_date_from").clear()
browser.find_element_by_name("birth_date_from").send_keys(timeStart)
browser.find_element_by_name("birth_date_to").clear()
browser.find_element_by_name("birth_date_to").send_keys(timeEnd)

browser.find_element_by_name("search").click()

bookName = time.strftime('%Y-%m-%d-%H-%M-%S',time.localtime(time.time()))
bookName = bookName+".xls"
startExcel(bookName)

for m in range(1,999,1):
	try:
		check(m,bookName)
		browser.find_element_by_link_text(u"下一页").click()
	except:
		print u"填写完毕"
		break
		








