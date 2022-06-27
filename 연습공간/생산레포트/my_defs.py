import seaborn as sns
import matplotlib.pyplot as plt


def season_sort(seasons):
    # 시즌 서열 나열
    o = {'SP':0, 'SU':1, 'FA':2, 'HO':3}
    # 결과로 쓸 리스트 생성
    

    # 가장 큰 시즌과 작은시즌 초기화
    min_s = 'HO99'
    max_s = 'SP01'
    # 시즌을 돌면서 검사
    for s in seasons:
        # 만약 년도와 시즌이 작다면
        if int(s[-2:]) < int(min_s[-2:]):
            min_s = s
        elif int(s[-2:]) == int(min_s[-2:]) and o[s[:2]] <= o[min_s[:-2]]:
            min_s = s
            
        if int(s[-2:]) > int(max_s[-2:]):
            max_s = s
        elif int(s[-2:]) == int(max_s[-2:]) and o[s[:2]] >= o[max_s[:-2]]:
            max_s = s
    return min_s, max_s

def date_func(d):
    y = d[6:]
    md = d[:-5]
    result = y + '/' + md
    return result.replace('/', '-')

def my_bar(my_df):
    plt.figure(figsize = (13,4))
    plt.rc('font', size=10)
    result = sns.countplot(x = "quote_state",data = my_df)
    return result

def my_pie(my_df, season):
    plt.rcParams["figure.figsize"] = (12, 9)
    plt.rc('font', size=15)
    t = my_df[my_df['lineplan_season']== season]
    
    total = len(t)
    requested = len(t[t['quote_state'] == 'Requested'])

    confirmed = len(t[t['quote_state'] == 'Confirmed'])

    the_others = total - requested - confirmed
    

    my_group = [requested, confirmed, the_others]
    my_label = ['Requested', 'Confirmed', 'In Progress']
    my_explode = (0.05, 0.05, 0.05)
    my_colors = ['#ff9999','#8fd9b6', '#ffc000']
    if the_others/total < 0.95 and requested/total < 0.95:    
        return plt.pie(my_group, autopct='%.1f%%', explode = my_explode, labels=my_label, shadow=True, startangle=260, colors=my_colors)
    else:
        return plt.pie(my_group, explode = my_explode, labels=['Requested','','The Others'], shadow=True, startangle=260, colors=my_colors)
    

def my_html():
    result = f'''
    <html>

    <body lang=KO style='tab-interval:40.0pt;word-wrap:break-word'>

    <div class=WordSection1>

    <p class=MsoNormal> Hi teams, This is Yeonggyun kim from DS costing </p>

    <p class=MsoNormal> This is an automated email for notice Sample Quote List is updated in Sephiroth  </p>
    <p class=MsoNormal> You don't have to respond to this email. </p>
    <p class=MsoNormal> You can refer to below tables and pictures.  </p>

    <img src="./pictures/220308_SU22-FA23/SU22_pieplot.png" alt="">
    <img src="./pictures/220308_SU22-FA23/FA22_pieplot.png" alt="">
    <img src="./pictures/220308_SU22-FA23/HO22_pieplot.png" alt="">
    <img src="./pictures/220308_SU22-FA23/SP23_pieplot.png" alt="">
    <img src="./pictures/220308_SU22-FA23/SU23_pieplot.png" alt="">
    <img src="./pictures/220308_SU22-FA23/FA23_pieplot.png" alt="">


    <p class=MsoNormal><span lang=EN-US><o:p>&nbsp;</o:p></span></p>

    <p class=MsoNormal>메일 <span class=GramE>보내기 입니다</span><span lang=EN-US>. </span></p>


    </div>

    </body>

    </html>
    '''
    return result



