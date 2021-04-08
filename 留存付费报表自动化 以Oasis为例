import pandas as pd
import numpy as np

# 此处可以修改输出的文件名
writer = pd.ExcelWriter('Oasis汇总数据0407_v5.xlsx')
'''
需要修改的参数：
q_path:问卷数据路径
t_path:tga数据路径
c_1和c_2:需要手动替换的字段
var:选取关心的变量
na_var:处理缺失值和替换字段的变量
i1_s_1和i1_p_1:单选题关心的变量
i1_s_1_1:i1_s_1的分类
v:多选题的选项
'''

# 1 处理问卷表，连接问卷表和tga表
q_path = 'C:\\Users\\74988\\Desktop\\python报表自动化测试\\测试数据\\Oasis问卷数据.xlsx'
t_path = 'C:\\Users\\74988\\Desktop\\python报表自动化测试\\测试数据\\Oasistga数据.xlsx'
q_data = pd.read_excel(q_path)
t_data = pd.read_excel(t_path)

q_question = list(q_data.columns)
q_answer = list(q_data.iloc[0,:])
q_title = []
for i in range(len(q_question)):
    q_title.append(q_answer[i]) if q_question[i][0:7] == 'Unnamed' else q_title.append(q_question[i])

'''
c_1为需要修改的字段的'列表'，c_2为需要修改字段和修改后字段的'字典'
此处FFA样例包括：（注：提问的问题需要替换成第一个选项，如第一印象的问题需要换成Graphics）
1.第一印象问题-->Graphics
2.游戏乐趣问题-->'Problem-solving (Cracking puzzles, etc.)'
3.尖叫度
4.推荐度
5.继续度
6.性别
7.年龄
8.职业...
请按顺序增加到c_1,c_2
'''
c_1 = [ 'From your first impression, are there any of these aspects attractive to you?',
        'Compared to similar games you\'ve played before, how would you rate this game?',
        'Based on your experience with this game, how likely are you to recommend it to your friends?',
        'Will you play more of this game?',
        'In the last year, which genre(s) of games have you played on Mobile/Tablet more than 1 month ?',
        'What is your gender?',
        'Please select your age range.',
        'What is your current occupation status?']
c_2 = {'From your first impression, are there any of these aspects attractive to you?':'Graphics',
        'Compared to similar games you\'ve played before, how would you rate this game?':'尖叫度',
        'Based on your experience with this game, how likely are you to recommend it to your friends?':'推荐度',
        'Will you play more of this game?':'继续度',
        'In the last year, which genre(s) of games have you played on Mobile/Tablet more than 1 month ?':'Gacha RPG',
        'What is your gender?': 'gender',
        'Please pick your age range.': 'age',
        'What is your current occupation?':'occupation'}
q_title = [c_2[i] if i in c_2 else i for i in q_title]
q_data.columns = q_title

#选取变量
var = [ 'uid',
        'Graphics',
        'Game setting',
        'Gameplay',
        '尖叫度',
        '推荐度',
        '继续度',
        'gender',
        'age',
        'occupation',
        'Gacha RPG',
        'Puzzle RPG',
        'Idle RPG',
        'Love Game',
        'Simulation',
        'Sports',
        'MOBA',
        'Strategy',
        'MMORPG',
        'Action',
        'Music',
        'Casual',
        'Gambling/Casino'
        ]

q_table = q_data.copy()[var].iloc[1:,:]
q_table = q_table.reset_index(drop=True)

# 缺失值处理和字段替换
na_var=['Graphics',
        'Game setting',
        'Gameplay',
        'Gacha RPG',
        'Puzzle RPG',
        'Idle RPG',
        'Love Game',
        'Simulation',
        'Sports',
        'MOBA',
        'Strategy',
        'MMORPG',
        'Action',
        'Music',
        'Casual',
        'Gambling/Casino']
for i in na_var: 
    q_table[i] = q_table[i].replace(i,1)
    q_table[i] = q_table[i].fillna(0)

# 处理核心玩家指标
q_table['尖叫度'] = q_table['尖叫度'].replace('Much better than similar games',5)
q_table['尖叫度'] = q_table['尖叫度'].replace('Almost the same as similar games',0)
q_table['尖叫度'] = q_table['尖叫度'].replace('Much worse than similar games',-5)

q_table['继续度'] = q_table['继续度'].replace('Definetely not',1)
q_table['继续度'] = q_table['继续度'].replace('Probably not',2)
q_table['继续度'] = q_table['继续度'].replace('It depends',3)
q_table['继续度'] = q_table['继续度'].replace('Probably will',4)
q_table['继续度'] = q_table['继续度'].replace('Definetely will',5)

q_table['推荐度'] = q_table['推荐度'].replace('Definitely will',10)
q_table['推荐度'] = q_table['推荐度'].replace('Definitely not',0)

# 增加新列：潜力玩家和核心玩家
n_1 = []
for i in range(len(q_table['Graphics'])):
    if int(q_table['Graphics'][i]+q_table['Game setting'][i]+q_table['Gameplay'][i]) <=0:
        n_1.append(-1)
    elif int(q_table['Graphics'][i]+q_table['Game setting'][i]+q_table['Gameplay'][i]) >= 3:
        n_1.append(1)
    else:
        n_1.append(0)
q_table['潜力玩家'] = pd.DataFrame(n_1)

n_2 = []
for i in range(len(q_table['尖叫度'])):
    if int(q_table['尖叫度'][i])>=3 and int(q_table['继续度'][i])>=4 and int(q_table['推荐度'][i])>=5:
        n_2.append(1)
    else:
        n_2.append(0)
q_table['核心玩家'] = pd.DataFrame(n_2)

q_t_data = pd.merge(q_table,t_data,on='uid')
q_table.to_excel(writer,'问卷数据',index=False)
t_data.to_excel(writer,'tga数据',index=False)
q_t_data.to_excel(writer,'问卷-tga数据',index=False)

# 2 计算tga数据和问卷数据的留存付费折算系数

# 留存折算系数
v_s_1 = ['uid','is_2r','is_3r','is_7r']
af_s_1 = {'uid':np.size,'is_2r':np.sum,'is_3r':np.sum,'is_7r':np.sum}
tga_survival_t = pd.pivot_table(t_data,index=['install_os'],values=v_s_1,aggfunc=af_s_1,margins=True)
q_t_survival_t = pd.pivot_table(q_t_data,index=['install_os'],values=v_s_1,aggfunc=af_s_1,margins=True)
survival_data = pd.DataFrame({'tga':list(tga_survival_t.iloc[-1,:]),'q':list(q_t_survival_t.iloc[-1,:])},index=tga_survival_t.columns)

survival_data = survival_data.T
a1 = survival_data['uid']
a2 = survival_data['is_2r']
a3 = survival_data['is_3r']
a7 = survival_data['is_7r']
s2 = a2/a1
s3 = a3/a1
s7 = a7/a1
survival_t=pd.DataFrame({'次日':list(s2),'三日':list(s3),'七日':list(s7)},index=['tga','问卷'])
survival_t = survival_t.T
tga_s = survival_t['tga']
q_s = survival_t['问卷']
s_index = tga_s/q_s
survival_index_t=pd.DataFrame({'tga留存率':list(tga_s),'问卷留存率':list(q_s),'折算系数':list(s_index)},index=['次日','三日','七日'])

# 付费折算系数
pay_value = {}
range_var = ['uid','_1pay_amount_cum','_2pay_amount_cum','_3pay_amount_cum','_4pay_amount_cum','_5pay_amount_cum','_6pay_amount_cum','_7pay_amount_cum','_8pay_amount_cum','_1ispay_cum','_2ispay_cum','_3ispay_cum','_4ispay_cum','_5ispay_cum','_6ispay_cum','_7ispay_cum','_8ispay_cum']
# 需要创造一个字典name将原名和中文名一一对应
chinese_name = ['uid','付费金额1','付费金额2','付费金额3','付费金额4','付费金额5','付费金额6','付费金额7','付费金额8','付费人数1','付费人数2','付费人数3','付费人数4','付费人数5','付费人数6','付费人数7','付费人数8']
name = {}
for i in range(len(range_var)):
    name[range_var[i]] = chinese_name[i]

for i in range(len(range_var)):
    if range_var[i] == 'uid':
        pay_value[range_var[i]] = np.size
    else:
        pay_value[range_var[i]] = np.sum

tga_pay_t = pd.pivot_table(t_data
                           ,index=['install_os']
                           ,values=range_var
                           ,aggfunc=pay_value
                           ,margins=True)
q_pay_t = pd.pivot_table(q_t_data
                           ,index=['install_os']
                           ,values=range_var
                           ,aggfunc=pay_value
                           ,margins=True)

tga_pay_t = tga_pay_t.T
q_pay_t = q_pay_t.T

new_name=[]
for i in list(tga_pay_t.index):
    new_name.append(new.get(i))

pay_data = pd.DataFrame({'tga':list(tga_pay_t.iloc[:,-1]),'q':list(q_pay_t.iloc[:,-1])},index=new_name)
pay_data.sort_index(inplace=True)

# 付费率
pay_date = ['1日','2日','3日','4日','5日','6日','7日','8日']
tga_p_rate=[]
q_p_rate=[]
for i in range(2):
    for j in range(1,9):
        if i == 0:          
            tga_p_rate.append(pay_data.iloc[j,i]/pay_data.iloc[0,i])
        else:
            q_p_rate.append(pay_data.iloc[j,i]/pay_data.iloc[0,i])
p_index=np.array(tga_p_rate)/np.array(q_p_rate)
pay_rate_t=pd.DataFrame({'tga付费率':tga_p_rate,'问卷付费率':q_p_rate,'折算系数':p_index},index=pay_date)

# ARPU
tga_arpu=[]
q_arpu=[]
for i in range(2):
    for j in range(9,len(pay_data)):
        if i == 0:          
            tga_arpu.append(pay_data.iloc[j,i]/pay_data.iloc[0,i])
        else:
            q_arpu.append(pay_data.iloc[j,i]/pay_data.iloc[0,i])
arpu_index=np.array(tga_arpu)/np.array(q_arpu)
arpu_index_t=pd.DataFrame({'tgaARPU':tga_arpu,'问卷ARPU':q_arpu,'折算系数':arpu_index},index=pay_date)

# ARPPU
tga_arppu=[]
q_arppu=[]
for i in range(2):
    for j in range(1,9):
        if i == 0:          
            tga_arppu.append(pay_data.iloc[j+8,i]/pay_data.iloc[j,i])
        else:
            q_arppu.append(pay_data.iloc[j+8,i]/pay_data.iloc[j,i])
arppu_index=np.array(tga_arppu)/np.array(q_arppu)
arppu_index_t=pd.DataFrame({'tgaARPU':tga_arppu,'问卷ARPU':q_arppu,'折算系数':arppu_index},index=pay_date)

pay_index_t = pd.concat([pay_rate_t,arpu_index_t,arppu_index_t],axis=1)

survival_index_t.to_excel(writer,'问卷-tga留存折算系数')
pay_index_t.to_excel(writer,'问卷-tga付费折算系数')

# 3 单选题分类（以潜力玩家分类）

# 潜力玩家数据留存透视表
i_s_1 = ['潜力玩家']
v_s_1 = ['uid','is_2r','is_3r','is_7r']
f_s_1 = {'uid':np.size,'is_2r':np.sum,'is_3r':np.sum,'is_7r':np.sum}
q_tga_t = pd.pivot_table(q_t_data,index=i_s_1,values=v_s_1,aggfunc=f_s_1)

# 留存率(已乘折算系数)
i_s_1_1 = ['非潜力玩家','低潜力玩家','高潜力玩家']
a1 = q_tga_t['uid']
a2 = q_tga_t['is_2r']
a3 = q_tga_t['is_3r']
a7 = q_tga_t['is_7r']
s2 = a2/a1*survival_index_t.iloc[0,2]
s3 = a3/a1*survival_index_t.iloc[1,2]
s7 = a7/a1*survival_index_t.iloc[2,2]
q_tga_survival_t = pd.DataFrame({'次日留存率':list(s2),'三日留存率':list(s3),'七日留存率':list(s7)},index=i_s_1_1)

#潜力玩家数据付费透视表
i_p_1 = ['潜力玩家']
f_p_1 = {}
for i in range(len(range_var)):
    if range_var[i] == 'uid':
        f_p_1[range_var[i]] = np.size
    else:
        f_p_1[range_var[i]] = np.sum
        
q_tga_pay_t = pd.pivot_table(q_t_data,index=i_p_1,values=range_var,aggfunc=f_p_1)
q_tga_pay_t.columns = new_name
q_tga_pay_t = q_tga_pay_t.T
q_tga_pay_t.sort_index(inplace=True)

# 付费率
q_tga_pay_rate={}
for i in range(1,9):
    q_tga_pay_rate[i] = np.array(q_tga_pay_t.iloc[i,:])/np.array(q_tga_pay_t.iloc[0,:])*pay_index_t.iloc[i-1,2]
q_tga_pay_rate_t = pd.DataFrame(q_tga_pay_rate,index=i_s_1_1)
q_tga_pay_rate_t.columns = pay_date

# ARPU
q_tga_arpu={}
for i in range(9,len(q_tga_pay_t)):
    q_tga_arpu[i] = np.array(q_tga_pay_t.iloc[i,:])/np.array(q_tga_pay_t.iloc[0,:])*pay_index_t.iloc[i-9,5]
q_tga_arpu_t = pd.DataFrame(q_tga_arpu,index=i_s_1_1)
q_tga_arpu_t.columns = pay_date

# ARPPU
q_tga_arppu={}
for i in range(1,9):
    q_tga_arppu[i] = np.array(q_tga_pay_t.iloc[i+8,:])/np.array(q_tga_pay_t.iloc[i,:])*pay_index_t.iloc[i-1,8]
q_tga_arppu_t = pd.DataFrame(q_tga_arppu,index=i_s_1_1)
q_tga_arppu_t.columns = pay_date

q_tga_survival_t.to_excel(writer,'潜力玩家留存率')
q_tga_pay_rate_t.to_excel(writer,'潜力玩家付费率')
q_tga_arpu_t.to_excel(writer,'潜力玩家arpu')
q_tga_arppu_t.to_excel(writer,'潜力玩家arppu')


# 4 多选题分类（以游戏乐趣分类）
v = [   'Gacha RPG',
        'Puzzle RPG',
        'Idle RPG',
        'Love Game',
        'Simulation',
        'Sports',
        'MOBA',
        'Strategy',
        'MMORPG',
        'Action',
        'Music',
        'Casual',
        'Gambling/Casino']

#留存透视表
q_tga_final = {}
# v_s_1 = ['uid','is_2r','is_3r','is_7r']
# f_s_1 = {'uid':np.size,'is_2r':np.sum,'is_3r':np.sum,'is_7r':np.sum}
for i in v:
    q_tga_t = pd.pivot_table(q_t_data,index=i,values=v_s_1,aggfunc=f_s_1)
    q_tga_final[i] = list(q_tga_t.iloc[-1,:])
q_tga_survival_t1 = pd.DataFrame(q_tga_final,index=v_s_1)
q_tga_survival_t1 = q_tga_survival_t1.T

#留存率(已乘系数)
a1 = q_tga_survival_t1['uid']
a2 = q_tga_survival_t1['is_2r']
a3 = q_tga_survival_t1['is_3r']
a7 = q_tga_survival_t1['is_7r']
s2 = a2/a1*survival_index_t.iloc[0,2]
s3 = a3/a1*survival_index_t.iloc[1,2]
s7 = a7/a1*survival_index_t.iloc[2,2]
q_tga_survival_t1 = pd.DataFrame({'次日留存率':list(s2),'三日留存率':list(s3),'七日留存率':list(s7)},index=v)

#付费透视表
q_tga_pay_final = {}
f_p = {}
for i in range(len(range_var)):
    if range_var[i] == 'uid':
        f_p[range_var[i]] = np.size
    else:
        f_p[range_var[i]] = np.sum
for i in v:
    q_tga_pay_t1 = pd.pivot_table(q_t_data,index=i,values=range_var,aggfunc=f_p)
    q_tga_pay_t1 = q_tga_pay_t1.T
    q_tga_pay_t1.sort_index(inplace=True)
    q_tga_pay_final[i] = list(q_tga_pay_t1.iloc[:,-1])

q_tga_pay_t1 = pd.DataFrame(q_tga_pay_final,index=new_name)
q_tga_pay_t1.sort_index(inplace=True)

# 付费率
q_tga_pay_rate1={}
for i in range(1,9):
    q_tga_pay_rate1[i] = np.array(q_tga_pay_t1.iloc[i,:])/np.array(q_tga_pay_t1.iloc[0,:])*pay_index_t.iloc[i-1,2]
q_tga_pay_rate_t1 = pd.DataFrame(q_tga_pay_rate1,index=v)
q_tga_pay_rate_t1.columns = pay_date

# ARPU
q_tga_arpu1={}
for i in range(9,len(q_tga_pay_t1)):
    q_tga_arpu1[i] = np.array(q_tga_pay_t1.iloc[i,:])/np.array(q_tga_pay_t1.iloc[0,:])*pay_index_t.iloc[i-9,5]
q_tga_arpu_t1 = pd.DataFrame(q_tga_arpu1,index=v)
q_tga_arpu_t1.columns = pay_date

# ARPPU
q_tga_arppu1={}
for i in range(1,9):
    q_tga_arppu1[i] = np.array(q_tga_pay_t1.iloc[i+8,:])/np.array(q_tga_pay_t1.iloc[i,:])*pay_index_t.iloc[i-1,8]
q_tga_arppu_t1 = pd.DataFrame(q_tga_arppu1,index=v)
q_tga_arppu_t1.columns = pay_date

q_tga_survival_t1.to_excel(writer,'游戏类型留存率')
q_tga_pay_rate_t1.to_excel(writer,'游戏类型付费率')
q_tga_arpu_t1.to_excel(writer,'游戏类型arpu')
q_tga_arppu_t1.to_excel(writer,'游戏类型arppu')

writer.save()
print('数据保存成功')
