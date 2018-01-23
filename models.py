# -*- coding: utf-8 -*-  
import os,docx,bs4,re,argparse,datetime,time,string,functools,sys
from bs4 import BeautifulSoup
from time import sleep
from docx import Document
from docx.shared import RGBColor
from docx.shared import Pt
import matplotlib.pyplot as plt
from functools import reduce
from init import init_para_cmplist
from docx.shared import Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.dml import MSO_THEME_COLOR
from imp import reload


################################################################################################################################################
# 设置文字的大小/粗细/颜色/斜体
# global document
document=Document()
# style=document.styles['Normal']
style1=document.styles['Normal']
paragraph= document.add_paragraph()
paragraph_format=paragraph.paragraph_format
font=style1.font
font.name="Calibri"
font.size = Pt(9)
font.italic = True
font.color.theme_color = MSO_THEME_COLOR.ACCENT_1
font.bold=True
paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
paragraph.space_after = Pt(2)
paragraph.space_before = Pt(2)
################################################################################################################################################
baseloca=os.path.dirname(os.getcwd())  #app同级目录
tmploca=os.path.join(baseloca,'templates')
resloca=os.path.join(baseloca,'result')
Month={'Jan':'01','Feb':'02','Mar':'03','Apr':'04','May':'05','June':'06','July':'07','Aug':'08','Sept':'09','Oct':'10','Nov':'11','Dec':'12'}
# testfile=os.path.join(tmploca,'awrrpt_1_30950_31013.html')
global file
################################################################################################################################################
# args[1]传入列的个数,args[2]传入th标签的内容,args[3]传入td标签内容,args[0]为document实例的，作用：生成表格
def generate_table(*args):
	table = args[0].add_table(rows=1, cols=args[1],style='Light Shading Accent 1')
	table._tblPr.autofit=True
	hdr_cells = table.rows[0].cells
	for i,j in enumerate(args[2]):
			if j==None:
				j=' '
			hdr_cells[i].text = str(j)
	for k in args[3]:
		row_cells = table.add_row().cells
		for i,j in enumerate(k):
			row_cells[i].text=str(j)


# 判断是否为不正常，假如是加红，然后写到最开始的表格中，第4个参数：1为小于参考值正常，0为大于参考值正常
@functools.lru_cache(maxsize=15)
def functionB(*args):
	row,col=args[2]
	try:
		if args[3]:
			if float(args[0]) > args[1] :
				run=document.tables[0].cell(row,col).paragraphs[0].add_run('不正常')
				run.font.color.rgb = RGBColor(0xFF, 0x00, 0x00)
			else:
				run=document.tables[0].cell(row,col).paragraphs[0].add_run('正常')
		else:
			if float(args[0]) < args[1] :
				run=document.tables[0].cell(row,col).paragraphs[0].add_run('不正常')
				run.font.color.rgb = RGBColor(0xFF, 0x00, 0x00)
			else:
				run=document.tables[0].cell(row,col).paragraphs[0].add_run('正常')		
	except:
		run=document.tables[0].cell(row,col).paragraphs[0].add_run('No Data found')
		run.font.color.rgb = RGBColor(0xFF, 0x00, 0x00)

# 显示mainreport中的内容
def MainReport():
	title=["AWR报告解析总结","AWR报告概况","主机资源概况","数据库内存配置","会话登录阶段","SQL解析阶段","SQL执行阶段","事务提交阶段","RAC Statistics","数据库参数建议"]
	document.add_heading("Main Report")
	for i in title:
		document.add_paragraph(i,style='List Bullet 2')

	# 一.AWR报告解析总结
	document.add_heading("一.AWR报告解析总结")
	title=['序号','检查流程','细项','检查结果']
	t1=['1','主机资源','CPU资源','']
	t2=[' ',' ','I/O资源','']
	t3=[' ',' ','内存资源','']
	t4=[' ',' ','网络资源','']
	t5=['2','数据库内存配置','Buffer cache','']
	t6=[' ',' ','Shared pool','']
	t7=[' ',' ','PGA','']
	t8=['3','会话登录阶段','会话连接时间','']
	t9=[' ',' ','登录次数','']
	t10=['4','SQL解析阶段','解析','']
	t11=[' ',' ','硬解析','']
	t12=['5','SQL执行阶段','执行时间','']
	t13=[' ',' ','逻辑读','']
	t14=[' ',' ','物理读','']
	t15=['6','事务提交阶段','提交响应时间','']
	t16=['7','RAC统计','集群响应时间','']
	t17=[' ',' ','节点内部通信性能','']
	t18=[' ',' ','节点内部通信PING延迟','']
	t19=['8','数据库参数建议','是否存在建议修改参数','']
	content=[t1,t2,t3,t4,t5,t6,t7,t8,t9,t10,t11,t12,t13,t14,t15,t16,t17,t18,t19]
	generate_table(document,4,title,content)

################################################################################################################################################
#第二部分：包括DB Name开始到Cursors/Session之间的内容,*title为每个表的标题栏，*result为表数据
################################################################################################################################################
def BasicSituation(*args):
	document.add_heading("二.AWR报告概况")
	soup=args[0].find_all("table",summary=re.compile("database instance information"))
	DBnametitle=[i.string.strip() for i in soup[0].find_all('th')]
	DBnameResult=[i.string.strip() for i in soup[0].find_all('td')]

	soup=args[0].find_all("table",summary=re.compile("host information"))
	HostNametitle=[i.string.strip() for i in soup[0].find_all('th')]
	HostNameResult=[i.string.strip() for i in soup[0].find_all('td')]

	soup=args[0].find_all("table",summary=re.compile("snapshot information"))
	tmp=[i.string.strip().replace(',','') for i in soup[0].find_all('td')]
	Snaptitle=[i.string for i in soup[0].find_all('th')]
	SnapResult=[tmp[i:i+len(Snaptitle)] for i in range(0,len(tmp),len(Snaptitle))]

	generate_table(document,len(DBnametitle),DBnametitle,[DBnameResult])
	document.add_paragraph(" ")
	generate_table(document,len(HostNametitle),HostNametitle,[HostNameResult])
	document.add_paragraph(" ")
	generate_table(document,len(Snaptitle),Snaptitle,SnapResult)

    #####################################  输出说明  #############################################################################3
	document.add_paragraph("输出说明:",style="Heading 9")
	tmp=dict(zip(DBnametitle,DBnameResult))
	db_name=tmp.get('DB Name','N/A')
	release=tmp.get('Release','N/A')
	rac='RAC' if (tmp.get('RAC','N/A') == 'YES') else '单机'
	document.add_paragraph("当前数据库名字为:%s，数据库版本为%s，为%s架构"%(db_name,release,rac),style="List Bullet 2")
	
	tmp=dict(zip(HostNametitle,HostNameResult))
	platform=tmp.get("Platform",'N/A')
	cpus=tmp.get("CPUs",'N/A')
	memory=tmp.get("Memory (GB)",'N/A')
	document.add_paragraph("当前数据库操作系统为:%s，有%s颗逻辑CPU，配置了%sGB内存"%(platform,cpus,memory),style="List Bullet 2")
	document.add_paragraph("数据库性能采样开始时间为:%s，结束时间为:%s，采样间隔为%s"%(SnapResult[0][2],SnapResult[1][2],SnapResult[2][2]),style="List Bullet 2")
	try:
		d_value=float(SnapResult[1][3])-float(SnapResult[0][3])
		if float(SnapResult[1][3])-float(SnapResult[0][3])<100 :
			tmp="会话连接数整体波动不大，总体表现正常" 
		else :
			"会话连接数量波动较大"
		document.add_paragraph("采样开始时,数据库连接数为:%s，采样结束时为%s，两者相差%s个，%s"%(SnapResult[0][3],SnapResult[1][3],d_value,tmp),style="List Bullet 2")
	except:
		pass

	try:
		time=SnapResult[3][2]
		cpus=int(float(time.replace('(mins)',''))/float(SnapResult[2][2].replace('(mins)','')))
		document.add_paragraph("数据库性能采样期间，DB Time消耗%s，消耗了%s颗逻辑CPU。"%(time,cpus),style="List Bullet 2")
	except:
		pass

################################################################################################################################################
#第三部分：主机资源
################################################################################################################################################	
def HostResource(*args):
	document.add_heading("三.主机资源概况")
	# 1.获取平均使用时间/平均空闲时间/平均空闲率 AVG_BUSY_TIME AVG_IDLE_TIME AVG_FREE
	soup=args[0].find_all("table",summary=re.compile("displays operating systems statistics"))
	try:
		tmp=soup[0].find_all('td')
		Statistic=[j.string for i,j in enumerate(tmp) if i%3==0]
		Value=[float(j.string.replace(',','')) for i,j in enumerate(tmp) if i%3==1]
		OperateSysStat=dict(zip(Statistic,Value))
		AVG_BUSY_TIME=round(OperateSysStat.get('AVG_BUSY_TIME')/60,3)
		AVG_IDLE_TIME=round(OperateSysStat.get('AVG_IDLE_TIME')/60,3)
		AVG_FREE=(round(float(AVG_IDLE_TIME)/float(AVG_BUSY_TIME)*100,3)) if (AVG_BUSY_TIME is not 'N/A' and AVG_IDLE_TIME is not 'N/A') else 'N/A'
	except:
		AVG_BUSY_TIME=AVG_IDLE_TIME=AVG_FREE='N/A'

	# 2.获取I/O资源信息，包括平均iops/平均每秒吞吐量/最大相应时间/平均响应时间  Total Requests=Total_req,Total (MB)=Total_MB
	soup=args[0].find_all("table",summary=re.compile("IO profile"))
	try:
		tmp=soup[0].find_all('td')
		IOProfilename=[j.string.replace(':','') for i,j in enumerate(tmp) if i%4==0]
		ReadWritePerSed=[j.string.replace(',','') for i,j in enumerate(tmp) if i%4==1]
		IOProfile=dict(zip(IOProfilename,ReadWritePerSed))
		Total_req=IOProfile.get('Total Requests')
		Total_MB=IOProfile.get('Total (MB)')
	except:
		Total_req=Total_MB='N/A'

	# 获得Av Rd(ms)的值 AvRd_MS
	soup=args[0].find_all("table",summary=re.compile("IO Statistics for different physical files"))
	try:
		tmp=[i.string.replace(' ','') for i in soup[0].find_all('th')]
		AvRdlocation=tmp.index('AvRd(ms)')
		AvRdTotal=[float(j.string.replace('\xa0','0').replace(',','')) for i,j in enumerate(soup[0].find_all('td')) if i%len(tmp)==AvRdlocation]
		AvRd_MS=max(AvRdTotal)
	except:
		AvRd_MS='N/A'

	# 获得平均响应时间，userio和systemio最大值 Avg_time
	soup=args[0].find_all("table",summary=re.compile("wait class statistics ordered by total wait time"))
	try:
		tmp=soup[0].find_all('td')
		WaitClass=[j.string for i,j in enumerate(tmp) if i%6==0]
		AvgWait=[j.string.replace(',','').replace('\xa0','0').replace('ms','') for i,j in enumerate(tmp) if i%6==3]
		l=dict(zip(WaitClass,AvgWait))
		Avg_time=max(float(l.get('User I/O',0)),float(l.get('System I/O',0)))
	except:
		Avg_time='N/A'
	
	#获得内存使用率 Menstatis
	soup=args[0].find_all("table",summary=re.compile("This table displays memory statistics"))
	try:
		tmp=soup[0].find_all('td')
		for i in tmp:
			if "% Host Mem used for SGA+PGA:" in i.string:
				Menstatis=max(float(i.next_sibling.string),float(i.next_sibling.next_sibling.string)) 
			else:
				pass
	except:
		Menstatis='N/A'

	# 获得每秒私网流量CacheLoadProfile
	soup=args[0].find_all("table",summary=re.compile("information about global cache load"))
	try:
		tmp=soup[0].find_all('td')
		Estd_traffic=[j.string.strip() for i,j in enumerate(tmp) if i%3==0]
		PerSedvalue=[j.string.replace(',','').replace('\xa0','0').strip() for i,j in enumerate(tmp) if i%3==1]
		l=dict(zip(Estd_traffic,PerSedvalue))
		CacheLoadProfile=l.get('Estd Interconnect traffic (KB)','N/A')
	except:
		CacheLoadProfile='N/A'

	# 判断是正常还是不正常
	functionB(AVG_FREE,10,(1,3),0)
	functionB(AvRd_MS,20,(2,3),1)
	functionB(Menstatis,90,(3,3),1)
	functionB(CacheLoadProfile,200480,(4,3),1)

	####################################### 表格 #######################################
	t0=['检查流程','细项','数值']
	t1=['CPU资源','平均使用时间(min)',AVG_BUSY_TIME]
	t2=['		','平均空闲时间(min)',AVG_IDLE_TIME]
	t3=['		','平均空闲率(%)',AVG_FREE]
	t4=['I/O资源','平均IOPS',Total_req.strip()]
	t5=['		','平均每秒吞吐量(MB)',Total_MB.strip()]
	t6=['		','最大响应时间(ms)',str(AvRd_MS)]
	t7=['		','平均响应时间(ms)',str(Avg_time)]
	t8=['内存资源','内存使用率(%)',str(Menstatis)]
	t9=['网络资源','每秒私网流量(KB)',CacheLoadProfile]
	generate_table(document,3,t0,[t1,t2,t3,t4,t5,t6,t7,t8,t9])
	document.add_paragraph("建议:",style="Heading 9")
	####################################### 建议 #######################################
	try:
		if float(AVG_FREE)<10 :
			tmp=("当前CPU平均空闲率为%s%%，"%(AVG_FREE),"空闲率低于10%，CPU资源紧张")
		else:
			tmp=("当前CPU平均空闲率为%s%%，"%(AVG_FREE),"空闲率高于10%，CPU资源正常")
		document.add_paragraph(tmp,style="List Bullet 2")
	except:
		pass

	try:
		if float(AvRd_MS)<20 :
			tmp="当前IO最大响应时间为%sms，低于参考值20ms，IO资源正常"%AvRd_MS
		elif float(AvRd_MS)<40:
			tmp="当前IO最大响应时间为%sms，高于参考值20ms，IO资源可能存在异常"%AvRd_MS
		else:
			tmp="当前IO最大响应时间为%sms，高于参考值20ms，IO资源存在异常"%AvRd_MS
		document.add_paragraph(tmp,style="List Bullet 2")
	except:
		pass

	try:
		if float(Menstatis)<80 :
			tmp=("当前内存使用率为:%s%%"%Menstatis,"，低于参考值80%，内存资源正常")
		else:
			tmp=("当前内存使用率为:%s%%"%Menstatis,"，高于参考值80%，内存资源不正常")
		document.add_paragraph(tmp,style="List Bullet 2")
	except:
		pass

	try:
		if float(CacheLoadProfile)>20480:
			tmp="当前私网流量为:%sKb/s，高于参考值20Mb/s，流量不正常"%CacheLoadProfile
		else:
			tmp="当前私网流量为:%sKb/s，低于参考值20Mb/s，流量正常"%CacheLoadProfile
		document.add_paragraph(tmp,style="List Bullet 2")
	except:
		pass

################################################################################################################################################
#第四部分：数据库内存配置建议PGA_USE_MB
################################################################################################################################################	
# @functools.lru_cache(maxsize=15)
# 获得的buffer pool中的值可能含有keep pool的值，所以进行筛选
def functionA(*args):
	flag=0
	l=args[0]
	for i,j in enumerate(l,1):
		try:
			if l[flag] < l[i]:
				flag+=1
			else:
				return l[:i]
		except:
			pass
	return l

def GetGraph(*args):
	table_summary=args[1]
	keyword=args[2]
	picname=os.path.basename(file).replace('docx','png')
	picloca=os.path.join(tmploca,picname)
	document.add_paragraph(args[3],style="Heading 9")
	try:
		soup=args[0].find_all("table",summary=re.compile(table_summary))
		aaa=soup[1] if len(soup)>1 else soup[0]
		t=[i.string.replace(' ','') for i in aaa.find_all('th')]
		total_len=len(t)
		start_len=t.index(keyword[0].replace(' ',''))
		end_len=t.index(keyword[1].replace(' ',''))
		tmp=aaa.find_all('td')
		x1=[float(j.string.replace(',','')) for i,j in enumerate(tmp) if i%total_len == start_len]
		y1=[float(j.string.replace(',','')) for i,j in enumerate(tmp) if i%total_len == end_len]

		tx1=functionA(x1)
		ty1=y1[:len(tx1)]
		# 绘制折线图
		plt.close()  #清图用，否则会和上一次的一起显示
		plt.plot(tx1,ty1,label='%s'%args[3],linewidth=3,color='r',marker='o',markerfacecolor='blue',markersize=4) 
		plt.xlabel(keyword[0]) 
		plt.ylabel(keyword[1]) 
		plt.title(args[3]) 
		plt.legend() 
		plt.savefig(picloca)
		# document.add_picture(picloca,width=Inches(1.25))
		document.add_picture(picloca)
		os.remove(picloca)	
	except IndexError:
		document.add_paragraph("No Data found",style="List Bullet 2")

# 将单位和数值分离，都转化为ms，float类型
def CovertUsToMs(argv):
	if 'ms' in argv:
		tmp=argv.split('ms')
		return (float(argv.split('ms')[0].replace(',','').strip()))
	elif 'us' in argv:
		return (float(argv.split('us')[0].replace(',','').strip())/1000)
	else:
		return float(argv)

# 第四部分：数据库内存配置
def MemoryConfig(*args):
	document.add_heading("四.数据库内存配置")
	################################  SGA取值  ################################
	soup=args[0].find_all("table",summary=re.compile('name and value of init.ora parameters'))
	try:
		for i in soup[0].find_all('td'):
			if "sga_target" in i.string.strip():
				SGA_VALUE=float(i.next_sibling.string.replace(',',''))
			else:
				pass
		sugg0="该数据库SGA内存管理方式为手动管理" if SGA_VALUE == 0.0 else "该数据库SGA内存管理方式为自动管理"
	except:
		SGA_VALUE='N/A'
		sugg0=''

	################################  BUFFER POOL取值  ################################
	soup=args[0].find_all("table",summary=re.compile("memory dynamic component statistics"))
	try:
		tmp=soup[0].find_all('td')
		comment=[j.string for i,j in enumerate(tmp) if i%7==0]
		maxsize=[j.string for i,j in enumerate(tmp) if i%7==1]
		l=dict(zip(comment,maxsize))
		DEFAULT_BUFFER_CACHE=l.get('DEFAULT buffer cache')
		KEEP_BUFFER_CACHE=l.get('KEEP buffer cache')
		RECYCLE_BUFFER_CACHE=l.get('RECYCLE buffer cache')
	except IndexError:
		RECYCLE_BUFFER_CACHE=DEFAULT_BUFFER_CACHE=KEEP_BUFFER_CACHE='N/A'

	soup=args[0].find_all("table",summary=re.compile("instance efficiency percentages"))
	try:
		tmp=soup[0].find_all('td')
		t=[j.string.replace(' ','') for i,j in enumerate(tmp) if i%2==0]
		tt=[j.string.strip() for i,j in enumerate(tmp) if i%2==1]
		l=dict(zip(t,tt))
		try:
			BUFFER_HIT=float(l.get('BufferHit%:'))
			LIBRARY_HIT=float(l.get('LibraryHit%:'))
		except ValueError:
			LIBRARY_HIT=BUFFER_HIT='N/A'
	except IndexError:
		LIBRARY_HIT=BUFFER_HIT='N/A'

	################################  建议  ################################
	sugg1="No Data found" if BUFFER_HIT == "N/A" else ("当前buffer cache命中率为:%s%%，"%BUFFER_HIT,'低于90%,不正常' if (BUFFER_HIT)<90 else "高于90%,正常")
	sugg2="No Data found"  if LIBRARY_HIT == "N/A" else ("当前shared pool命中率为:%s%%，"%LIBRARY_HIT,'低于98%,不正常' if (LIBRARY_HIT)<98 else "高于98%,正常")	
	sugg3="No Data found"  if LIBRARY_HIT == "N/A" else ("当前keep pool命中率为:%s%%，"%BUFFER_HIT,'低于99%,不正常' if (BUFFER_HIT)<99 else "高于99%,正常")

	functionB((BUFFER_HIT),90,(5,3),0)
	functionB((LIBRARY_HIT),98,(6,3),0)

	################################  PGA取值  ################################
	soup=args[0].find_all("table",summary=re.compile("memory dynamic component statistics"))
	try:
		for i in soup[0].find_all('td'):
			if "PGA Target" in i.string:
				try:
					PGA_USE=float(i.next_sibling.next_sibling.next_sibling.next_sibling.string.replace(',','').strip())
				except:
					PGA_USE='N/A'
			else:
				pass
	except:
		PGA_USE='N/A'
	try:
		soup=args[0].find_all("table",summary=re.compile("shared pool advisory. Size factor, estimated library cache size"))
		# length=len(soup[0].find_all('th'))
		tmp=soup[0].find_all('td')
		t=[j.string.strip() for i,j in enumerate(tmp) if i%8==1]
		tt=[j.string.replace(',','').strip() for i,j in enumerate(tmp) if i%8==2]
		t1=dict(zip(t,tt))
		EST_LC_SIZE=t1.get('1.00')
	except:
		EST_LC_SIZE='N/A'

	################################  表格  ###################################
	document.add_heading('1.数据库内存资源',level=2)
	t1=['内存组件','细项','数值(MB)','命中率(%)']
	t0=['SGA','/',SGA_VALUE,'N/A']
	t2=['Buffer cache','Default Pool',DEFAULT_BUFFER_CACHE,BUFFER_HIT]
	t3=['		','Keep Pool',KEEP_BUFFER_CACHE,BUFFER_HIT]
	t4=['		','Recycle Pool',RECYCLE_BUFFER_CACHE,'N/A']
	t5=['Shared Pool','Library Cache',EST_LC_SIZE,LIBRARY_HIT]
	t6=['		','Dictionary Cache','N/A','N/A']
	t7=['PGA','/',PGA_USE,'N/A']
	generate_table(document,4,t1,[t0,t2,t3,t4,t5,t6,t7])
	document.add_paragraph("建议:",style="Heading 9")
	if sugg0:
		document.add_paragraph(sugg0,style="List Bullet 2")
	else:
		pass
	document.add_paragraph(sugg1,style="List Bullet 2")
	document.add_paragraph(sugg2,style="List Bullet 2")
	document.add_paragraph(sugg3,style="List Bullet 2")

	###################################  内存抖动信息  ##########################################
	soup=args[0].find_all("table",summary=re.compile("memory dynamic component statistics. Begin snap size, current size, min size"))
	document.add_heading('2.SGA内存抖动信息',level=2)
	try:
		tmp=[i.string for i in soup[0].find_all('td')]
		y=[tmp[i:i+7] for i in range(0,len(tmp),7)]
		t1=['内存组件','开始大小(MB)','当前大小(MB)','最小值(MB)','最大值(MB)','操作次数','操作类型']
		generate_table(document,7,t1,y)
		comp_list=[i[0] for i in y]
		opercount_list=[float(i[5]) for i in y]
		maxsize_list=[float(i[4].replace(',','')) for i in y]
		document.add_paragraph("建议:",style="Heading 9")
		tmp=[]
		try:
			for i,j in enumerate(opercount_list):
				if not(j == 0):
					tmp.append(maxsize_list[i])
				else:
					pass
			if tmp:
				document.add_paragraph("操作次数一分钟内平均超过1次,建议设置内存手动管理,参考建议值为%s"%max(tmp),style="List Bullet 2")
				document.add_paragraph("SGA内存配置存在抖动现象,内存抖动对于高响应需求的OLTP业务影响较大，需要避免，可以根据内存配置曲线或咨询DBA进行调整",style="List Bullet 2")
			else:
				document.add_paragraph("SGA没有发生抖动，表现正常",style="List Bullet 2")
		except:
			pass
	except:
		document.add_paragraph("No Data found",style="List Bullet 2")

	###################################  PGA溢出信息  ##########################################
	soup=args[0].find_all("table",summary=re.compile("PGA aggregate target histograms"))
	document.add_heading('3.PGA溢出信息',level=2)
	try:
		flag=0
		title=[i.string.strip() for i in soup[0].find_all('th')]
		length=len(title)
		tmp=[i.string.replace(',','').strip() for i in soup[0].find_all('td')]
		value=[tmp[i:i+length] for i in range(0,len(tmp),length)]
		t=[i[-2:] for i in value]
		generate_table(document,len(title),title,value)
		for j in t:
			if j.count('0') < 2:
				flag+=1
			else:
				pass
		document.add_paragraph("建议:",style="Heading 9")
		if flag:
			document.add_paragraph('1-pass Execs或m-pass Execs出现的次数为:%s'%flag,style="List Bullet 2")
			document.add_paragraph("当前PGA配置为:%sMb，PGA内存配置可能不足"%PGA_USE,style="List Bullet 2")
			run=document.tables[0].cell(7,3).paragraphs[0].add_run('不正常')
			run.font.color.rgb = RGBColor(0xFF, 0x00, 0x00)
		else:
			document.add_paragraph('PGA无溢出',style="List Bullet 2")
			document.tables[0].cell(7,3).text='正常'
	except:
		document.add_paragraph('No Data found',style="List Bullet 2")
		run=document.tables[0].cell(7,3).paragraphs[0].add_run('No Data found')
		run.font.color.rgb = RGBColor(0xFF, 0x00, 0x00)

	###################################  内存配置建议 ##########################################
	document.add_heading('4.内存配置建议',level=2)
	GetGraph(args[0],'MTTR advisory',['Size for Est (M)','Estimated Phys Reads (thousands)'],'Buffer Pool Advisory')
	GetGraph(args[0],'shared pool advisory',['Shared Pool Size(M)','Est LC Load Time (s)'],'Shared Pool Advisory')
	GetGraph(args[0],'PGA memory advisory',['PGA Target Est (MB)','Estd PGA Overalloc Count'],'PGA memory advisory')
	GetGraph(args[0],'SGA target advisory for different SGA target sizes',['SGA Target Size (M)','Est Physical Reads'],'SGA Target Advisory')
	GetGraph(args[0],'Streams Pool Advisory',['Size for Est (MB)','Est Spill Time (s)'],'Streams Pool Advisory')
	document.add_paragraph("建议:",style="Heading 9")
	document.add_paragraph("内存建议曲线一边观察横轴变化与纵轴变化占比，找出拐点，衡量突变点投入产出比，比如增加少量buffer cache值，可以大幅减少物理读，则建议调整，如果增加大量buffer cache值，只能略微减少物理读，则不建议调整",style="List Bullet 2")


# 第五部分，会话登陆阶段
def TimeModelStat(*args):
	# 获取connection management call elapsed time的times和% of DB Time值(TIMES/DB_TIME_TIMEMODEL)
	soup=args[0].find_all("table",summary=re.compile("time model statistics"))
	try:
		for i in soup[0].find_all('td'):
			if "connection management call elapsed time" in i.string: 
				TIMES=i.next_sibling.string
				DB_TIME_TIMEMODEL=i.next_sibling.next_sibling.string
			else:
				pass
	except:
		TIMES=DB_TIME_TIMEMODEL="N/A"

	# 获取Load Profile的Logons的Per Second和Per Transaction值
	soup=args[0].find_all("table",summary=re.compile("load profile"))
	try:
		for i in soup[0].find_all('td'):
			if "Logons:" in i.string: 
				try:
					LOGONS_PER_SEC=float(i.next_sibling.string)
					LOGONS_PER_TRAN=float(i.next_sibling.next_sibling.string)
				except ValueError:
					LOGONS_PER_SEC=LOGONS_PER_TRAN="N/A"
			else:
				pass
	except:
		LOGONS_PER_SEC=LOGONS_PER_TRAN="N/A"

	t1=["Statistic Name","Time (s)",r"% of DB Time"]
	t2=["connection management call elapsed time",TIMES,DB_TIME_TIMEMODEL]
	t3=["指标","Per Second","Per Transaction"]
	t4=["Logons",LOGONS_PER_SEC,LOGONS_PER_TRAN]
	# awr报告解析总结 会话登录阶段 会话连接时间的结果
	try:
		TimeModelSug="当前会话连接时间占DB Time比例为%s%%,"%DB_TIME_TIMEMODEL
		if (float(DB_TIME_TIMEMODEL)>1) :
			TimeModelSug+="超过参考值1%，会话连接性能不正常。建议检查数据库连接配置情况，如密码错误，短连接，防火墙，登录访问策略等"
		else:
			TimeModelSug+="低于参考值1%,会话连接性能正常"
	except:
		TimeModelSug='N/A'

	# awr报告解析总结 会话登录阶段 登录次数的结果
	try:
		LogonsSug="当前会话登录每秒为%s个,"%LOGONS_PER_SEC
		if (float(LOGONS_PER_SEC)>80):
			LogonsSug+="超过参考值80个，登录应用连接数量不正常 ，建议尽量减少登录频率，建议使用长连接"
		else:
			LogonsSug+="少于参考值每秒80个，其每秒登录数在监听处理范围内"
	except:
		LogonsSug = "N/A"

	functionB(DB_TIME_TIMEMODEL,1,(8,3),1)
	functionB(LOGONS_PER_SEC,80,(9,3),1)

	document.add_heading("五.会话登录阶段")
	document.add_heading('1.时间模型',level=2)
	generate_table(document,3,t1,[t2])
	document.add_paragraph("建议:",style="Heading 9")
	document.add_paragraph(TimeModelSug,style="List Bullet 2")
	document.add_heading("2.指标",level=2)
	generate_table(document,3,t3,[t4])
	document.add_paragraph("建议:",style="Heading 9")
	document.add_paragraph(LogonsSug,style="List Bullet 2")


# 第六部分，sql解析阶段
def TimeModel(*args):
	# 1.时间模型,每行的值存在timemodelstatlist中
	document.add_heading("六.SQL解析阶段")
	document.add_heading('1.时间模型',level=2)	
	soup=args[0].find_all("table",summary=re.compile("time model statistics"))
	try:
		timemodelstatlist=[]
		contextlist=["parse time elapsed","hard parse elapsed time","failed parse elapsed time","hard parse (sharing criteria) elapsed time",\
					"hard parse (bind mismatch) elapsed time"]
		t1=['Statistic Name','Time (s)',r'% of DB Time']
		for i in soup[0].find_all('td'):
			if i.string.strip().replace('\n','') in contextlist: 
				tmplist=[]
				tmplist.append(i.string.replace('\n',''))
				tmplist.append(i.next_sibling.string.replace(',',''))
				tmplist.append(i.next_sibling.next_sibling.string)	
				timemodelstatlist.append(tmplist)
			else:
				pass
		Hard_Parse=float(timemodelstatlist[1][1])
		Hard_Fail_Parse=(reduce(lambda x,y:x+y,[float(i[1]) for i in timemodelstatlist if i[0] in ['hard parse elapsed time','failed parse elapsed time']] ))
		Parse_Time=float(timemodelstatlist[0][1])
		Parse_DBtime=float(timemodelstatlist[0][2])
		try:
			sugg1="当前失败解析与硬解析占比解析时间为:%s"%round(100*Hard_Fail_Parse/Parse_Time,3)
			if 100*(Hard_Fail_Parse/Parse_Time)<10 :
				sugg1+="%，小于建议值10%，SQL解析正常"
			else:
				sugg1+="%，大于建议值10%，SQL解析存在问题，建议检查是否存在异常等待事件，业务sql是否使用绑定变量，是否存在高版本SQL"
		except:
			sugg1="parse time为0"

		try:
			sugg2="解析时间占比DB Time占比%s%%"%Parse_DBtime
			if Parse_DBtime > 20 :
				sugg2+="，超过参考值20%，不正常"
			else:
				sugg2+="，低于参考值20%，正常"
		except:
			sugg2='N/A'
		generate_table(document,3,t1,timemodelstatlist)
		document.add_paragraph("建议:",style="Heading 9")
		document.add_paragraph(sugg1,style="List Bullet 2")
		document.add_paragraph(sugg2,style="List Bullet 2")
	except:
		document.add_paragraph("No Data found")

	# 2.指标，每行的值存在loadproflist中
	soup=args[0].find_all("table",summary=re.compile("load profile"))
	loadproflist=[]
	t2=['指标','Per Second','Per Transaction','ms/次']
	contextlist2=["Parses","Hard parses"]
	for i in soup[0].find_all('td'):
		if i.string.replace('\n','').replace('(SQL):','').replace(':',"").strip() in contextlist2: 
			tmplist=[]
			tmplist.append(i.string.replace('\n',''))
			tmplist.append(i.next_sibling.string.strip().replace(",",""))
			tmplist.append(i.next_sibling.next_sibling.string.strip())
			loadproflist.append(tmplist)		
		else:
			pass
	document.add_heading("2.指标",level=2)

	try:
		loadproflist[0].append(round((1000*Parse_Time)/float(loadproflist[0][1]),3))
	except:
		loadproflist[0].append('N/A')
	try:
		loadproflist[1].append(round((1000*Hard_Parse)/float(loadproflist[1][1]),3))
	except:
		loadproflist[1].append("N/A")

	# sql解析阶段  解析的值
	try:
		sugg2="当前解析为：%s ms/次%s"%(loadproflist[0][3],("，低于参考值2ms/次，解析正常" if loadproflist[0][3] < 2 else "，高于参考值2ms/次，解析偏慢"))
	except:
		sugg2="数据有问题"

	#sql解析阶段  硬解析的值
	try:
		sugg3="当前硬解析为:%s ms/次%s"%(loadproflist[1][3],("，低于参考值5ms/次，硬解析正常" if loadproflist[1][3] < 5 else "，高于参考值5ms/次，硬解析偏慢"))
	except:
		sugg3="数据有问题"

	functionB(loadproflist[0][3],2,(10,3),1)
	functionB(loadproflist[1][3],5,(11,3),1)

	generate_table(document,4,t2,loadproflist)
	document.add_paragraph("建议:",style="Heading 9")
	document.add_paragraph(sugg2,style="List Bullet 2")
	document.add_paragraph(sugg3,style="List Bullet 2")

	# 3.异常等待事件
	soup=args[0].find_all("table",summary=re.compile("Foreground Wait Events and their wait statistics"))
	soup1=args[0].find_all("table",summary=re.compile("background wait events statistics"))
	document.add_heading("3.主要等待事件",level=2)
	# fore_wait_events=[]
	t3=["等待事件","Waits","Total Wait Time (sec)","% DB time","Wait Avg(ms)"]
	contextlist3=["library cache load lock","library cache lock","library cache pin","library cache: mutex S","library cache: mutex X","row cache lock","cursor: mutex S","cursor: mutex X","cursor: pin S wait on X","cursor: pin S","cursor: pin X","latch: row cache objects","latch: shared pool"]
	tmp2=[]
	tmp1=[]
	try:
		# Foreground Wait Events
		for i in soup[0].find_all('td'):
			if i.string.strip().replace('\n','') in contextlist3: 
				wait_event=i.string.replace('\n','')
				wait_waits=float(i.next_sibling.string.strip().replace(',',''))
				wait_totaltime=float(i.next_sibling.next_sibling.next_sibling.string.strip().replace(',',''))
				wait_dbtime=float(i.next_sibling.next_sibling.next_sibling.next_sibling.next_sibling.next_sibling.string.strip().replace(',',''))
				tmp2.append([wait_event,wait_waits,wait_totaltime,wait_dbtime])
			else:
				pass
		#  background wait events
		for j in soup1[0].find_all('td'):
			if j.string.strip().replace('\n','') in contextlist3: 
				wait_event=j.string.replace('\n','')
				wait_waits=float(j.next_sibling.string.strip().replace(',',''))
				wait_totaltime=float(j.next_sibling.next_sibling.next_sibling.string.strip().replace(',',''))
				try:
					wait_dbtime=float(i.next_sibling.next_sibling.next_sibling.next_sibling.next_sibling.next_sibling.string.strip().replace(',',''))
				except:
					pass
				tmp1.append([wait_event,wait_waits,wait_totaltime,wait_dbtime])
			else:
				pass

		# tmp1中若存在和tmp2，则和tmp2中的值进行相加，否则自己获得平均相应时间之后加到tmp2中
		for i in tmp1:
			k=1
			try:
				for j in tmp2:
					if i[0] in j:
						j[1]+=i[1]
						j[2]+=i[2]
						K=0
					else:
						pass
				if k:
					tmp2.append(i)
			except:
				pass
		# 进行排序和筛选，取total wait time前五的值
		tmp=[]
		fore_wait_events=(sorted(tmp2,key=lambda x:x[2],reverse=True)[:5])
		for i in fore_wait_events:
			i.append(round(1000*float(i[2])/float(i[1]),3))
			tmp.append(round(1000*float(i[2])/float(i[1]),3))

		hard_parse_list=['latch: row cache objects','latch: shared pool','row cache lock','cursor: mutex X','library cache: mutex X','cursor: pin X','library cache lock']
		soft_parse_list=['library cache: mutex S','cursor: pin S','cursor: mutex S']
		soft_soft_list=['cursor: pin S']
		
		generate_table(document,5,t3,fore_wait_events)

		t1=[]
		t2=[]
		t3=[]
		document.add_paragraph("建议:",style="Heading 9")
		for i in fore_wait_events:
			if i[4] > 1:
				if i[0] in hard_parse_list:
					t1.append(i[0])
				elif i[0] in soft_parse_list:
					t2.append(i[0])
				elif i[0] in soft_soft_list:
					t3.append(i[0])
				document.add_paragraph("当前%s的单次响应时间为:%sms，超过参考值1ms"%(i[0],i[4]),style="List Bullet 2")	
			else:
				pass
		if t1:
			document.add_paragraph("其中"+'/'.join(t1)+"一般是由于SQL未使用绑定变量，多版本，执行计划异常导致",style="List Bullet 2")
		else:
			pass
		if t2:
			document.add_paragraph("其中"+'/'.join(t1)+"一般可以增加session_cached_cursor数量，或使用游标缓存来优化",style="List Bullet 2")
		else:
			pass
		if t3:
			document.add_paragraph("其中"+'/'.join(t1)+"一般软软解析是属于比较理想的解析类型，如果出现软软解析竞争，则说明系统的并发量较大，建议控制业务并发量，辅助可以使用游标缓存或HINTS来优化",style="List Bullet 2")
		else:
			pass			
	except:
		document.add_paragraph("No Data found")

	# 4.解析数最高的SQL
	soup=args[0].find_all("table",summary=re.compile("top SQL by number of parse calls"))
	document.add_heading("4.解析数最高的SQL",level=2)
	try:
		length=len(soup[0].find_all('th'))
		tmp=[i.string.replace(',','') for i in soup[0].find_all('td')]
		t5=["Parse Calls","Executions","% Total Parses","SQL Id"]
	
		tmp1=[tmp[i:i+4] for i in range(0,len(tmp),length)]
		SQL_ParseCall=sorted(tmp1,key=lambda x:float(x[0]),reverse=True)[:5]
		cursor_cache=[]
		for i in SQL_ParseCall[:5]:
			try:
				cursor_cache.append(round(100*float(i[0])/float(i[1]),3))
			except:
				cursor_cache.append(0)

		generate_table(document,4,t5,SQL_ParseCall)
		document.add_paragraph("建议:",style="Heading 9")
		document.add_paragraph("TOP SQL的解析调用为:%s次，SQL ID为%s，SQL Text为%s"%(SQL_ParseCall[0][1],SQL_ParseCall[0][3],tmp[length-1]),style="List Bullet 2")
		if max(cursor_cache) > 50:
			document.add_paragraph("检查高解析SQL，存在Parse Calls/Executions大于50%，未良好使用游标缓存功能(游标缓存：oracle建议在中间层延缓关闭常用游标时间，此游标再次被执行时，不需要解析阶段，可直接绑定执行)",style="List Bullet 2")
		else:
			pass
	except:
		document.add_paragraph("No Data found")	

	# 5.版本数最高的SQL
	soup=args[0].find_all("table",summary=re.compile("top SQL by version counts"))
	document.add_heading("5.版本数最高的SQL",level=2)
	try:
		length=len(soup[0].find_all('th'))
		tmp=[i.string.replace(',','') for i in soup[0].find_all('td')]
		t5=["Version count","Executions","SQL Id"]

		tmp1=[tmp[i:i+3] for i in range(0,len(tmp),length)]
		SQL_VersionCount=sorted(tmp1,key=lambda x:float(x[0]),reverse=True)[:5]
		
		generate_table(document,3,t5,SQL_VersionCount)
		document.add_paragraph("建议:",style="Heading 9")
		document.add_paragraph("高版本会导致SQL解析效率下降，影响性能，通常是由于绑定变量MISMATCH或者oracle 11G ACS自适应游标共享新特性导致，具体原因可以从V$SQL_SHARED_CURSOR视图中查看高版本原因。",style="List Bullet 2")
		document.add_paragraph("针对绑定变量MISMATCH，一般可以在应用端调大绑定变量所占内存空间解决",style="List Bullet 2")
		document.add_paragraph("针对ACS，一般可以直接关闭新特性解决。关闭方法:",style="List Bullet 2")
		document.add_paragraph('alter system set "_optimizer_adaptive_cursor_sharing"=false scope=both;',style="List Bullet 3")
		document.add_paragraph('alter system set "_optimizer_extended_cursor_sharing_rel"=none  scope=both;',style="List Bullet 3")
	except:
		document.add_paragraph("No Data found")	

# 七.SQL执行阶段
def SqlExecuteTime(*args):
	# 1.第一部分，时间模型
	soup=args[0].find_all("table",summary=re.compile("time model statistics"))
	document.add_heading("七.SQL执行阶段")
	document.add_heading('1.时间模型',level=2)
	try:
		t1=["Statistic Name","Time (s)",r"% of DB Time"]
		SQL_EXECUTE_TIME=[]
		for i in soup[0].find_all('td'):
			if "sql execute elapsed time" in i.string:
				SQL_EXECUTE_TIME.append(i.string)
				SQL_EXECUTE_TIME.append(i.next_sibling.string.strip().replace(',',''))
				SQL_EXECUTE_TIME.append(i.next_sibling.next_sibling.string.strip().replace(',',''))
			else:
				pass

		generate_table(document,3,t1,[SQL_EXECUTE_TIME])

		document.add_paragraph("建议:",style="Heading 9")

		sug1="当前SQL执行时间占DB Time比例为%s%%"%SQL_EXECUTE_TIME[2]
		if float(SQL_EXECUTE_TIME[2])>80:
			sug1 += "，大于80%,SQL执行正常"
		else:
			sug1 += "，小于参考值80%，SQL执行不正常，总体性能可能会比较差"

		document.add_paragraph(sug1,style="List Bullet 2")
		functionB(SQL_EXECUTE_TIME[2],80,(12,3),0)
	except:
		document.add_paragraph("No Data found")	

	# 2.指标
	soup=args[0].find_all("table",summary=re.compile("load profile"))
	t2=['指标','Per Second','Per Transaction']
	document.add_heading("2.指标",level=2)
	SQL_EXEC_LIST=[]
	contextlist2=["Executes","Logical read","Physical read","Physical write","Read IO","Write IO","Logical reads","Physical reads","Physical writes"]
	for i in soup[0].find_all('td'):
		if i.string.replace('(SQL)','').replace(':',"").replace('(blocks)','').strip() in contextlist2: 
			tmplist=[]
			tmplist.append(i.string.replace('\n',' ').strip())
			tmplist.append(i.next_sibling.string.strip().replace(',',''))
			tmplist.append(i.next_sibling.next_sibling.string.strip().replace(',',''))	
			SQL_EXEC_LIST.append(tmplist)
		elif i.string.replace(':',"")in 'DB CPU(s):':
			DB_Time=float(i.next_sibling.string.strip().replace(',','').replace('\n',''))
		else:
			pass	
	t=sorted(SQL_EXEC_LIST)

	generate_table(document,3,t2,SQL_EXEC_LIST)
	document.add_paragraph("建议:",style="Heading 9")
	try:
		tmp=round(1000*DB_Time/float(t[1][1]),3)
		sug1="当前每次逻辑读的响应时间为%s，%s"%(tmp,("超过建议值5ms，逻辑读不正常" if tmp > 5 else "低于建议值5ms，逻辑读正常"))
		functionB(tmp,5,(13,3),1)
	except:
		sug1 = "逻辑读数据存在问题"

	try:
		tmp=round(1000*DB_Time/float(t[2][1]),3)
		sug2="当前每次物理读的响应时间为%sms，%s"%(tmp,("超过建议值10ms，物理读不正常" if tmp > 10 else "低于建议值10ms，物理读正常"))
		functionB(tmp,10,(14,3),1)
	except:
		sug2 = "物理读数据存在问题"		

	
	

	document.add_paragraph(sug1,style="List Bullet 2")
	document.add_paragraph(sug2,style="List Bullet 2")

	# 3.异常等待事件
	soup=args[0].find_all("table",summary=re.compile("Foreground Wait Events and their wait statistics"))
	document.add_heading("3.主要等待事件",level=2)
	try:
		tmp=soup[0].find_all('td')
		t3=["等待事件","Waits","Total Wait Time (sec)","Wait Avg(ms)","% DB time"]
		contextlist3=["latch: cache buffers chains","latch: checkpoint queue latch","buffer busy waits","read by other session","db file sequential read","db file scattered read"]
		tmp2=[]
		for i in tmp:
			if i.string.strip().replace('\n','') in contextlist3: 
				tmplist=[]
				tmplist.append(i.string.replace('\n',''))
				tmplist.append(float(i.next_sibling.string.strip().replace(',','')))
				tmplist.append(float(i.next_sibling.next_sibling.next_sibling.string.strip().replace(',','')))		
				tmplist.append(CovertUsToMs(i.next_sibling.next_sibling.next_sibling.next_sibling.string))
				tmplist.append(float(i.next_sibling.next_sibling.next_sibling.next_sibling.next_sibling.next_sibling.string.strip().replace(',','')))
				tmp2.append(tmplist)
			else:
				pass

		fore_wait_events=(sorted(tmp2,key=lambda x:x[2],reverse=True)[:5])
		# print(fore_wait_events)
		generate_table(document,5,t3,fore_wait_events)
		document.add_paragraph("建议:",style="Heading 9")
		for i in fore_wait_events:
			if "latch: cache buffers chains" in i:
				document.add_paragraph("%s:一般由于逻辑读过大导致，往往是由于未有合理索引机制或者执行计划变化导致SQL走全表扫描"%i[0],style="List Bullet 2")
			elif "buffer busy waits" in i:
				document.add_paragraph("%s:DML和DML或者DML和SELECT之间的竞争导致，一般是由于批量DML或者执行计划变更导致SQL效率下降导致"%i[0],style="List Bullet 2")
			elif "read by other session" in i:
				document.add_paragraph("%s:是由于热点块冲突，一般是由于执行计划变更导致SQL执行效率下降导致"%i[0],style="List Bullet 2")
			elif "db file sequential read" in i:
				document.add_paragraph("%s:如果此等待事件响应超过20ms，则存储性能可能存在问题或者SQL执行计划变动导致物理读增加"%i[0],style="List Bullet 2")
			elif "db file scattered read" in i:
				document.add_paragraph("%s:如果此等待事件响应超过20ms，则存储性能可能存在问题或者SQL执行计划变动导致物理读增加"%i[0],style="List Bullet 2")
	except:
		document.add_paragraph("No Data found")	

	# 4.执行时间最长的SQL
	soup=args[0].find_all("table",summary=re.compile("top SQL by elapsed time"))
	document.add_heading("4.执行事件最长的SQL",level=2)
	try:
		length=len(soup[0].find_all('th'))
		tmp=[i.string.strip().replace(',','') for i in soup[0].find_all('td')]
		SQL_ELAPSED_TIME=[tmp[i:i+7] for i in range(0,len(tmp),length)]
		t4=["Elapsed Time (s)","Executions","Elapsed Time per Exec (s)","%Total","%CPU",r"%IO","SQL Id"]
		generate_table(document,7,t4,SQL_ELAPSED_TIME[:5])
		document.add_paragraph("建议:",style="Heading 9")
		document.add_paragraph("TOP SQL的执行时间为:%s(s)，TOP SQL ID为:%s，TOP SQL TEXT为:%s"%(SQL_ELAPSED_TIME[0][0],SQL_ELAPSED_TIME[0][-1],tmp[length-1]),style="List Bullet 2")	
		document.add_paragraph("检查总执行时间高及单次执行时间高的SQL，从索引机制，业务逻辑层面优化SQL，降低执行时间",style="List Bullet 2")	
	except:
		document.add_paragraph("No Data found")	

	# 5.执行次数最多的SQL
	soup=args[0].find_all("table",summary=re.compile("top SQL by number of executions"))
	document.add_heading("5.执行次数最多的SQL",level=2)
	try:
		length=len(soup[0].find_all('th'))
		tmp=[i.string.strip().replace(',','') for i in soup[0].find_all('td')]
		SQL_Executions=[tmp[i:i+7] for i in range(0,len(tmp),length)]
		t5=["Executions","Rows Processed","Rows per Exec","Elapsed Time (s)","%CPU",r"%IO","SQL Id"]
		# print("TOP SQL的执行时间为%s(s),SQL ID为%s"%(SQL_Executions[0][0],SQL_Executions[0][-1]))
		generate_table(document,7,t5,SQL_Executions[:5])
		document.add_paragraph("建议:",style="Heading 9")
		document.add_paragraph("TOP SQL每秒执行次数为%s，执行频率过高容易引起热点块争用，SQL ID为:%s，SQL TEXT为:%s"%(\
								SQL_Executions[0][0],SQL_Executions[0][-1],tmp[length-1]),style="List Bullet 2")				
		document.add_paragraph("检查执行次数高的SQL，优化SQL业务逻辑",style="List Bullet 2")	
	except:
		document.add_paragraph("No Data found")

	# 6.消耗CPU时间最多的SQL
	soup=args[0].find_all("table",summary=re.compile("top SQL by CPU time"))
	document.add_heading("6.消耗CPU时间最多的SQL",level=2)
	try:
		length=len(soup[0].find_all('th'))
		tmp=[i.string.strip().replace(',','') for i in soup[0].find_all('td')]
		t=[tmp[i:i+8] for i in range(0,len(tmp),length)]
		t6=["CPU Time (s)","Executions","Rows Processed","Rows per Exec","%CPU",r"%IO","SQL Id"]
		SQL_CPUtime=[i[:4]+i[5:8] for i in t]
		# print(tmp[length-1])
		generate_table(document,7,t6,SQL_CPUtime[:5])
		document.add_paragraph("建议:",style="Heading 9")
		document.add_paragraph("TOP SQL的CPU消耗时间为:%s(s)，TOP SQL ID为:%s，TOP SQL TEXT为:%s"%(SQL_CPUtime[0][0],SQL_CPUtime[0][-1],tmp[length-1]),style="List Bullet 2")	
		document.add_paragraph("检查CPU消耗高的SQL，优化SQL，减少单次执行消耗CPU。建议检查索引机制是否合理，统计信息是否准确，SQL业务逻辑是否可优化",style="List Bullet 2")
	except:
		document.add_paragraph("No Data found")	

	# 7.消耗逻辑读最多的SQL
	soup=args[0].find_all("table",summary=re.compile("top SQL by buffer gets"))
	document.add_heading("7.消耗逻辑读最多的SQL",level=2)
	try:
		length=len(soup[0].find_all('th'))
		tmp=[i.string.strip().replace(',','') for i in soup[0].find_all('td')]
		t=[tmp[i:i+8] for i in range(0,len(tmp),length)]
		t7=["Buffer Gets","Executions","Gets per Exec","%Total","%CPU",r"%IO","SQL Id"]
		
		SQL_LOG=[i[:4]+i[5:8] for i in t]
		generate_table(document,7,t7,SQL_LOG[:5])
		document.add_paragraph("建议:",style="Heading 9")
		document.add_paragraph("TOP SQL的逻辑读为:%s，TOP SQL ID为:%s，TOP SQL TEXT为:%s"%(SQL_LOG[0][0],SQL_LOG[0][-1],tmp[length-1]),style="List Bullet 2")

		document.add_paragraph("检查逻辑读高的SQL，优化SQL，减少单次执行消耗逻辑读。逻辑读高的SQL，通常意味着执行计划不合理，也会消耗更多的CPU，建议检查索引机制是否合理，统计信息是否准确，SQL业务逻辑是否可优化",style="List Bullet 2")	
	except:
		document.add_paragraph("No Data found")	

	# 8.消耗物理读最多的SQL
	soup=args[0].find_all("table",summary=re.compile("top SQL by physical reads"))
	document.add_heading("8.消耗物理读最多的SQL",level=2)
	try:
		length=len(soup[0].find_all('th'))
		tmp=[i.string.strip().replace(',','') for i in soup[0].find_all('td')]
		t=[tmp[i:i+8] for i in range(0,len(tmp),length)]
		t8=["Physical Reads","Executions","Reads per Exec","%Total","%CPU",r"%IO","SQL Id"]
		SQL_PHY=[i[:4]+i[5:8] for i in t]
		generate_table(document,7,t8,SQL_PHY[:5])
		document.add_paragraph("建议:",style="Heading 9")
		document.add_paragraph("TOP SQL的物理读为:%s，TOP SQL ID为:%s，TOP SQL TEXT为:%s"%(SQL_PHY[0][0],SQL_PHY[0][-1],tmp[length-1]),style="List Bullet 2")
		document.add_paragraph("检查物理读高的SQL，优化SQL，减少单次执行消耗物理读。物理读高往往对应大量的全表扫描，建议检查索引机制是否合理，统计信息是否准确，SQL业务逻辑是否可优化。对于频度不高的高物理读消耗SQL，建议选择业务空闲时间段执行。对于不可避免的高物理读消耗SQL，可以通过将表格KEEP到内存中来减少物理读",style="List Bullet 2")	
	except:
		document.add_paragraph("No Data found")	

	# 9.逻辑读最多的对象
	soup=args[0].find_all("table",summary=re.compile("top segments by logical reads"))
	document.add_heading("9.逻辑读最多的对象",level=2)
	try:
		length=len(soup[0].find_all('th'))
		tmp=[i.string.strip().replace(',','') for i in soup[0].find_all('td')]
		SQL_LOGICAL_RS=[tmp[i:i+7] for i in range(0,len(tmp),length)]
		t9=["Owner","Tablespace Name","Object Name","Subobject Name","Obj. Type",r"Logical Reads","%Total"]
		generate_table(document,7,t9,SQL_LOGICAL_RS[:5])
		document.add_paragraph("建议:",style="Heading 9")
		document.add_paragraph("逻辑读最多的用户为:%s，表空间名为:%s，对象名为:%s，对象类型为:%s，逻辑读为%s"%(SQL_LOGICAL_RS[0][0],SQL_LOGICAL_RS[0][1],SQL_LOGICAL_RS[0][2],SQL_LOGICAL_RS[0][4],SQL_LOGICAL_RS[0][5]),style="List Bullet 2")
		document.add_paragraph("检查逻辑读高的对象，优化对应SQL及应用，降低其逻辑读数量，可以结合逻辑读高的SQL进行分析",style="List Bullet 2")	
	except:
		document.add_paragraph("No Data found")	

	# 10.物理读最多的对象
	soup=args[0].find_all("table",summary=re.compile("top segments by physical reads"))
	document.add_heading("10.物理读最多的对象",level=2)
	try:
		length=len(soup[0].find_all('th'))
		tmp=[i.string.strip().replace(',','') for i in soup[0].find_all('td')]
		SQL_physical_RS=[tmp[i:i+7] for i in range(0,len(tmp),length)]
		t10=["Owner","Tablespace Name","Object Name","Subobject Name","Obj. Type","Physical Reads","%Total"]
		generate_table(document,7,t10,SQL_physical_RS[:5])
		document.add_paragraph("建议:",style="Heading 9")
		document.add_paragraph("物理读最多的用户为:%s，表空间名为:%s，对象名为:%s，对象类型为:%s，物理读为%s"%(SQL_physical_RS[0][0],SQL_physical_RS[0][1],SQL_physical_RS[0][2],SQL_physical_RS[0][4],SQL_physical_RS[0][5]),style="List Bullet 2")
		document.add_paragraph("检查物理读高的对象，优化SQL及业务，降低或避免物理读，可以结合物理读高的SQL进行分析。如果操作系统有多余的内存及比较重组的cpu资源，则可以根据实际情况将其keep到内存中。",style="List Bullet 2")	
	except:
		document.add_paragraph("No Data found")	

	# 11.直接路径读最多的对象
	soup=args[0].find_all("table",summary=re.compile("top segments by direct physical reads"))
	document.add_heading("11.直接路径读最多的对象",level=2)
	try:
		length=len(soup[0].find_all('th'))
		tmp=[i.string.strip().replace(',','') for i in soup[0].find_all('td')]
		SQL_Dphysical_RS=[tmp[i:i+7] for i in range(0,len(tmp),length)]
		t11=["Owner","Tablespace Name","Object Name","Subobject Name","Obj. Type","Direct Reads","%Total"]
		generate_table(document,7,t11,SQL_Dphysical_RS[:5])
		document.add_paragraph("建议:",style="Heading 9")
		document.add_paragraph("物理读最多的用户为:%s，表空间名为:%s，对象名为:%s，对象类型为:%s，直接路径读读为%s"%(SQL_Dphysical_RS[0][0],SQL_Dphysical_RS[0][1],SQL_Dphysical_RS[0][2],SQL_Dphysical_RS[0][4],SQL_Dphysical_RS[0][5]),style="List Bullet 2")
		document.add_paragraph("检查直接路径读高的对象，直接路径路往往对应并行查询，全表扫描操作，优化相关SQL及应用，降低资源使用率，判断直接路径读参数是否关闭，关闭直接路径读的方法:",style="List Bullet 2")	
		document.add_paragraph('alter system set "_serial_direct_read"=never scope=both;',style="List Bullet 3")
	except:
		document.add_paragraph("No Data found")	


# 第八部分:事务提交阶段
def TransCommit(*args):
	# 1.第一部分，时间模型
	soup=args[0].find_all("table",summary=re.compile("wait class statistics ordered by total wait time"))
	document.add_heading("八.事务提交阶段")
	document.add_heading('1.时间模型',level=2)
	try:
		length=len(soup[0].find_all('th'))
		tmp=[i.string.strip().replace(',','') for i in soup[0].find_all('td')]
		Tran_Commit=[]
		t=[tmp[i:i+5] for i in range(0,len(tmp),length)]
		for i in t:
			if "Commit" in i:
				Tran_Commit.extend(i)
			else:
				pass
		t1=["Wait Class","Waits","Total Wait Time (sec)","Avg Wait (ms)",r"% of DB Time"]

		sugg1="当前提交响应时间为:%sms"%CovertUsToMs(Tran_Commit[3])
		if CovertUsToMs(Tran_Commit[3]) > 10 :
			sugg1 += "提交响应时间操作超过参考值10ms，提交响应时间不正常，提交问题一般由于IO响应慢，日志成员过小，组数过少，日志量异常增大，并发异常增大导致" 
		else :
			sugg1 += "提交响应时间操作低于参考值10ms，提交响应时间正常"

		functionB(CovertUsToMs(Tran_Commit[3]),10,(15,3),1)

		generate_table(document,5,t1,[Tran_Commit])
		document.add_paragraph("建议:",style="Heading 9")
		document.add_paragraph("%s"%sugg1,style="List Bullet 2")
	except:
		document.add_paragraph("No Data found",style="List Bullet 2")

	# 2.指标,用于获取Redo size (bytes)和Transactions值
	soup=args[0].find_all("table",summary=re.compile("load profile"))
	length=len(soup[0].find_all('th'))
	tmp=[i.string.strip().replace(',','').replace('\n','') for i in soup[0].find_all('td')]
	t2=['指标','Per Second','Per Transaction']
	t=[tmp[i:i+3] for i in range(0,len(tmp),length)]
	LordProfile=[]
	for i in t:
		if "Redo size" in i[0]:
			LordProfile.append(i)
		elif "Transactions" in i[0]:
			LordProfile.append(i)
			tran_per=i[1]
		elif "DB CPU(s):" in i:
			DB_CPU=i[1]
		else:
			pass

	# 获得user commits和user rollbacks值
	soup=args[0].find_all("table",summary=re.compile("Key Instance activity statistics"))
	try:
		for i in soup[0].find_all('td'):
			if i.string.strip().replace('\n','') in ["user commits","user rollbacks"]:
				t=[]
				t.append(i.string)
				t.append(i.next_sibling.next_sibling.string.replace(',',''))
				t.append(i.next_sibling.next_sibling.next_sibling.string.replace(',',''))
				LordProfile.append(t)
			else:
				pass	
	except:
		pass
	document.add_heading("2.指标",level=2)
	generate_table(document,3,t2,LordProfile)

	try:
		per_tran_time=round(float(DB_CPU)/float(tran_per),3)
	except:
		per_tran_time='N/A'

	document.add_paragraph("建议:",style="Heading 9")
	document.add_paragraph("关注回滚次数，如果回滚比例过高，说明业务逻辑不合理",style="List Bullet 2")
	document.add_paragraph("平均事务响应时间为%ss"%per_tran_time,style="List Bullet 2")
	# 3.异常等待事件
	soup=args[0].find_all("table",summary=re.compile("Foreground Wait Events and their wait statistics"))
	soup1=args[0].find_all("table",summary=re.compile("background wait events statistics"))
	document.add_heading("3.主要等待事件",level=2)
	try:
		t3=["等待事件","Waits","Total Wait Time (sec)","Wait Avg(ms)",""]
		contextlist3=["latch: redo allocation","latch: redo copy","latch: redo writing","log file sync"]
		fore_wait_events=[]
		for i in soup[0].find_all('td'):
			if i.string.strip().replace('\n','') in contextlist3: 
				tmplist=[]
				tmplist.append(i.string.replace('\n',''))
				tmplist.append(float(i.next_sibling.string.strip().replace(',','')))
				tmplist.append(float(i.next_sibling.next_sibling.next_sibling.string.strip().replace(',','')))	
				tmplist.append(CovertUsToMs(i.next_sibling.next_sibling.next_sibling.next_sibling.string))
				tmplist.append(i.next_sibling.next_sibling.next_sibling.next_sibling.next_sibling.next_sibling.string.strip())
				fore_wait_events.append(tmplist)
			else:
				pass
		for i in soup1[0].find_all('td'):
			if "log file parallel write" in i.string.strip():
				tmplist=[]
				tmplist.append(i.string.replace('\n',''))
				tmplist.append(float(i.next_sibling.string.strip().replace(',','')))
				tmplist.append(float(i.next_sibling.next_sibling.next_sibling.string.strip().replace(',','')))	
				tmplist.append(CovertUsToMs(i.next_sibling.next_sibling.next_sibling.next_sibling.string))
				tmplist.append(i.next_sibling.next_sibling.next_sibling.next_sibling.next_sibling.next_sibling.string.strip())
				fore_wait_events.append(tmplist)
			else:
				pass				

		generate_table(document,5,t3,fore_wait_events)
		document.add_paragraph("建议:",style="Heading 9")
		for i in fore_wait_events:
			sugg="当前%s的平均响应时间为%sms，"%(i[0],i[3])
			if i[3] > 10:
				sugg += "超过参考值10ms，日志写存在问题"
			else:
				sugg += "低于参考值10ms，日志写正常"
			document.add_paragraph(sugg,style="List Bullet 2")
		document.add_paragraph("若log file sync与log file parallel write如果相差过大，则说明并发提交量过大",style="List Bullet 2")
	except:
		document.add_paragraph("No Data found")	

	# 4.行级锁最多的对象
	soup=args[0].find_all("table",summary=re.compile("top segments by row lock waits"))
	document.add_heading("4.行级锁最多的对象",level=2)
	try:
		tmp=[i.string.strip().replace(',','') for i in soup[0].find_all('td')]
		length=len(soup[0].find_all('th'))
		ROW_LOCK_WAITS=[tmp[i:i+7] for i in range(0,len(tmp),length)]

		t4=["Owner","Tablespace Name","Object Name","Subobject Name","Obj. Type","Row Lock Waits",r"% of Capture"]
		generate_table(document,7,t4,ROW_LOCK_WAITS[:5])
		document.add_paragraph("建议:",style="Heading 9")
		document.add_paragraph("行级锁最多的用户为:%s,表空间名为:%s，对象名为:%s，对象类型为:%s，行级锁的数量为%s"%(\
							ROW_LOCK_WAITS[0][0],ROW_LOCK_WAITS[0][1],ROW_LOCK_WAITS[0][2],ROW_LOCK_WAITS[0][4],ROW_LOCK_WAITS[0][5]),style="List Bullet 2")
		document.add_paragraph("检查行锁最多的对象及对应SQL，检查SQL效率，并发量及业务逻辑，减少单次操作的锁定时间，加快事务提交时间",style="List Bullet 2")	
	except:
		document.add_paragraph("No Data found")	

# 第九部分RAC Statistics
def RacTotalWaitTime(*args):
	# 1.第一部分，时间模型
	soup=args[0].find_all("table",summary=re.compile("wait class statistics ordered by total wait time"))
	document.add_heading("九.RAC Statistics")
	document.add_heading('1.时间模型',level=2)	
	try:
		tmp=[i.string.strip().replace(',','') for i in soup[0].find_all('td')]
		RAC_Cluster=[]
		t=[tmp[i:i+5] for i in range(0,len(tmp),6)]
		for i in t:
			if "Cluster" in i:
				RAC_Cluster=i
			else:
				pass
		t1=["Wait Class","Waits","Total Wait Time (sec)","Avg Wait (ms)",r"% of DB Time"]

		sugg1="当前平均等待时间为:%sms,"%RAC_Cluster[3]
		if  CovertUsToMs(RAC_Cluster[3]) > 5:
			sugg1 += "超过参考值5ms，集群间响应存在问题，检查集群内部通信网络问题或应用数据访问节点间交互频度"
		else :
			sugg1 += "低于参考值5ms，集群内部通信网络正常"

		functionB(CovertUsToMs(RAC_Cluster[3]),5,(16,3),5)

		generate_table(document,5,t1,[RAC_Cluster])
		document.add_paragraph("建议:",style="Heading 9")
		document.add_paragraph("%s"%sugg1,style="List Bullet 2")	 
	except:
		document.add_paragraph("No Data found",style="List Bullet 2")

	# 2.指标
	soup=args[0].find_all("table",summary=re.compile("global cache load"))
	document.add_heading("2.指标",level=2)
	try:
		tmp=[i.string.strip().replace(',','').replace('\n','') for i in soup[0].find_all('td')]
		t2=['指标                               ','Per Second','Per Transaction']
		contextlist2=["Global Cache blocks received:","Global Cache blocks served:","GCS/GES messages received:","DBWR Fusion writes:","Estd Interconnect traffic (KB)"]
		t=[tmp[i:i+3] for i in range(0,len(tmp),3)]
		LordProfile=[]
		for i in t:
			if i[0].strip() in contextlist2:
				LordProfile.append(i)
			else:
				pass

		generate_table(document,3,t2,LordProfile)
		document.add_paragraph("建议:",style="Heading 9")
		document.add_paragraph("当前网络流量为%sKb/s"%sorted(LordProfile)[1][1],style="List Bullet 2")	
	except IndexError:
		document.add_paragraph("No Data found",style="List Bullet 2")

	# 3.Global Cache and Enqueue Services
	soup=args[0].find_all("table",summary=re.compile("workload characteristics for global"))
	document.add_heading("3.Global Cache and Enqueue Services",level=2)
	try:
		tmp=[i.string.replace(',','').replace('\n','').strip() for i in soup[0].find_all('td')]
		contextlist3=["Avg global cache cr block receive time (ms):","Avg global cache current block receive time (ms):","Avg global cache cr block build time (ms):",\
						"Avg global cache cr block send time (ms):","Global cache log flushes for cr blocks served %:","Avg global cache cr block flush time (ms):",\
						"Avg global cache current block pin time (ms):","Avg global cache current block send time (ms):","Global cache log flushes for current blocks served %:",\
						"Avg global cache current block flush time (ms):"]
		t=[tmp[i:i+2] for i in range(0,len(tmp),2)]
		t3=[]
		Global_Cache=[]
		for i in t:
			if i[0].strip() in contextlist3:
				Global_Cache.append(i)
				t3.extend(i)
			else:
				pass
		t1=[j for i,j in enumerate(t3) if i%2==0]
		t2=[]
		for j in [j for i,j in enumerate(t3) if i%2==1]:
			try:
				t2.append(float(j))
			except:
				t2.append(0.0)
		t4=dict(zip(t1,t2))
		del t4['Global cache log flushes for current blocks served %:']
		del t4['Global cache log flushes for cr blocks served %:']
		sugg3="no Data found" if not(t4.values()) else ("平均响应时间超过20ms,节点内部通信性能存在问题,交互太频繁" if max(t4.values())>20 else "节点内部通信性能正常")

		generate_table(document,2,['Global Cache and Enqueue Services','数值'],Global_Cache)
		document.add_paragraph("建议:",style="Heading 9")

		flag=0
		for i in Global_Cache:
			if '%' not in i[0]:
				try:
					if 'Avg global cache current block pin time (ms):' in i:
						if float(i[1]) > 8:
							sugg3 ="当前%s的值为%sms，超过参考值8ms，不正常"%(i[0].replace('(ms)','').replace(':',''),i[1])
							flag=1
							document.add_paragraph("%s"%sugg3,style="List Bullet 2")	
						else:
							pass
					else:
						if CovertUsToMs(i[1]) > 4:
							sugg3 ="当前%s的值为%sms，超过参考值4ms，不正常"%(i[0].replace('(ms)','').replace(':',''),i[1])
							flag=1
							document.add_paragraph("%s"%sugg3,style="List Bullet 2")	
						else:
							pass
				except:
					pass
			else:
				pass
			
		if flag:
			document.add_paragraph("节点内部通信性能不正常",style="List Bullet 2")
			run=document.tables[0].cell(17,3).paragraphs[0].add_run('不正常')
			run.font.color.rgb = RGBColor(0xFF, 0x00, 0x00)
		else:
			document.add_paragraph("节点内部通信性能正常",style="List Bullet 2")
			document.tables[0].cell(17,3).text = '正常'
	except:
		document.add_paragraph("No Data found",style="List Bullet 2")


	# 4.节点内部通信PING延迟
	soup=args[0].find_all("table",summary=re.compile(" IC ping latency statistics"))
	document.add_heading("4.节点内部通信PING延迟",level=2)
	try:
		title=[i.string.strip() for i in soup[0].find_all('th')]
		length=len(title)
		tmp=[i.string.replace(',','').replace('\n','').strip() for i in soup[0].find_all('td')]
		value=[tmp[i:i+length] for i in range(0,len(tmp),length)]
		generate_table(document,length,title,value)

		document.add_paragraph("建议:",style="Heading 9")
		if float(value[0][2]) < 1 and float(value[0][5]) < 1 and float(value[1][2]) <1 and float(value[1][5]) <1:
			document.add_paragraph("节点内部通信PING延迟低于参考值1ms，正常",style="List Bullet 2")	
			document.tables[0].cell(18,3).text = '正常'
		else:
			document.add_paragraph("节点内部通信PING延迟超过参考值1ms，不正常",style="List Bullet 2")	
			run=document.tables[0].cell(18,3).paragraphs[0].add_run('不正常')
			run.font.color.rgb = RGBColor(0xFF, 0x00, 0x00)
	except:
		document.add_paragraph("No Data found",style="List Bullet 2")


	# 5.Dynamic Remastering Stats 
	soup=args[0].find_all("table",summary=re.compile("Dynamic Remastering Stats. . times are in seconds"))
	document.add_heading("5.Dynamic Remastering Stats",level=2)	
	try:
		title=[i.string.strip() for i in soup[0].find_all('th')]
		length=len(title)
		tmp=[i.string.replace(',','').replace('\n','').strip() for i in soup[0].find_all('td')]
		value=[tmp[i:i+length] for i in range(0,len(tmp),length)]
		generate_table(document,length,title,value)	
		document.add_paragraph("建议:",style="Heading 9")
		document.add_paragraph("建议关闭drm,关闭方法:",style="List Bullet 2")
		document.add_paragraph("""alter system set "_gc_policy_time"=0 scope=spfile sid='*';""",style="List Bullet 3")
	except:
		document.add_paragraph("建议:",style="Heading 9")	
		document.add_paragraph("无",style="List Bullet 2")

	# 6.异常等待事件
	soup=args[0].find_all("table",summary=re.compile("Foreground Wait Events and their wait statistics"))
	document.add_heading("6.主要等待事件",level=2)
	try:
		tmp=soup[0].find_all('td')
		t4=["等待事件","Waits","Total Wait Time (sec)","Wait Avg(ms)",'% DB time']
		contextlist3=["gc cr multi block request","gc buffer busy acquire","gc current block busy","gc cr block busy","gcs log flush sync","gc current multi block request"]
		fore_wait_events=[]
		for i in tmp:
			if i.string.strip().replace('\n','') in contextlist3: 
				tmplist=[]
				tmplist.append(i.string.replace('\n',''))
				tmplist.append(float(i.next_sibling.string.strip().replace(',','')))
				tmplist.append(float(i.next_sibling.next_sibling.next_sibling.string.strip().replace(',','')))	
				tmplist.append(CovertUsToMs(i.next_sibling.next_sibling.next_sibling.next_sibling.string))
				tmplist.append(i.next_sibling.next_sibling.next_sibling.next_sibling.next_sibling.next_sibling.string.strip())
				fore_wait_events.append(tmplist)
			else:
				pass

		generate_table(document,5,t4,fore_wait_events)
		mm=[float(i[3]) for i in fore_wait_events]
		if max(mm) > 1 :
			sugg4="平均响应时间超过参考值1ms，集群节点间网络通信存在问题或者集群节点交互过于频繁，建议检查节点间通信网络，减少节点间交互"
		else:
			sugg4="平均响应时间低于参考值1ms，集群节点间网络通信正常"

		document.add_paragraph("建议:",style="Heading 9")
		document.add_paragraph(sugg4,style="List Bullet 2")	
	except:
		document.add_paragraph("No Data found",style="List Bullet 2")

	# 7.Global Cache Buffer Busy最多的对象
	soup=args[0].find_all("table",summary=re.compile("top segments by row lock waits"))
	document.add_heading("5.Global Cache Buffer Busy最多的对象",level=2)
	try:
		length=len(soup[0].find_all('th'))
		tmp=[i.string.strip().replace(',','') for i in soup[0].find_all('td')]
		GLOBAL_CACHE_BUFFER_BUSY=[tmp[i:i+7] for i in range(0,len(tmp),length)]
		t5=["Owner","Tablespace Name","Object Name","Subobject Name","Obj. Type","GC Buffer Busy",r"% of Capture"]
		generate_table(document,7,t5,GLOBAL_CACHE_BUFFER_BUSY[:5])
		document.add_paragraph("建议:",style="Heading 9")
		document.add_paragraph("Global Cache Buffer Busy最多的用户为%s，表空间为%s，对象类型为%s，对象类型为%s，GC Buffer Busy为%s"%(GLOBAL_CACHE_BUFFER_BUSY[0][0],\
								GLOBAL_CACHE_BUFFER_BUSY[0][1],GLOBAL_CACHE_BUFFER_BUSY[0][2],GLOBAL_CACHE_BUFFER_BUSY[0][4],GLOBAL_CACHE_BUFFER_BUSY[0][-2]),style="List Bullet 2")
		document.add_paragraph("检查节点间交互对象情况，检查相关SQL及业务，尽量减少节点间对象交互",style="List Bullet 2")	
	except:
		document.add_paragraph("No Data found",style="List Bullet 2")
	
def InitParaAdvise(*args):
	soup=args[0].find_all("table",summary=re.compile("name and value of init.ora parameters"))
	document.add_heading("十.数据库参数建议")
	try:
		advise_list=[]
		diff_list=[]
		tmp=[i.string.strip() for i in soup[0].find_all('td')]
		t1=[j for i,j in enumerate(tmp) if i%3==0]
		t2=[j for i,j in enumerate(tmp) if i%3==1]
		t3=dict(zip(t1,t2))
		key_word=init_para_cmplist.keys()
		for i in key_word:
			t=t3.get(i,'N/A')
			advise_list.append([i,t,init_para_cmplist[i]])
			if t is not 'N/A' and not(t == init_para_cmplist[i]):
				diff_list.append([i,init_para_cmplist[i]])
			else:
				pass

		generate_table(document,3,['参数','当前值','建议值'],advise_list)
		document.add_paragraph("建议:",style="Heading 9")
		if diff_list:
			for i in diff_list:
				document.add_paragraph("建议修改%s的值为%s:"%(i[0],i[1]),style="List Bullet 2")
			document.tables[0].cell(19,3).text = '存在'
		else:
			document.add_paragraph('无',style="List Bullet 2")
			document.tables[0].cell(19,3).text = '不存在'

	except:
		document.add_paragraph("No Data found",style="List Bullet 2")

def run(*args):
	try:
		soup=BeautifulSoup(open(args[0],'r',encoding='utf-8'),'html.parser')
	except :
		soup=BeautifulSoup(open(args[0],'r',encoding='gbk'),'html.parser')
	global file
	file=os.path.join(resloca,os.path.basename(args[0]).replace('html','docx'))
	MainReport()
	BasicSituation(soup)
	HostResource(soup)
	MemoryConfig(soup)
	TimeModelStat(soup)
	TimeModel(soup)
	SqlExecuteTime(soup)
	TransCommit(soup)
	RacTotalWaitTime(soup)
	InitParaAdvise(soup)
	document.save(file)



def test():
	# testfile=os.path.join(tmploca,'awrrpt_1_46376_46378.html')
	testfile=os.path.join(tmploca,'awrrpt_1_209199_209200.html')
	run(testfile)	



if __name__ == "__main__":
	test()

	# run(str(sys.argv[1]))	