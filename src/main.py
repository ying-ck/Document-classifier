from sklearn.feature_extraction.text import TfidfVectorizer
from sklearn.model_selection import train_test_split
from sklearn.naive_bayes import MultinomialNB
from sklearn.metrics import accuracy_score
from tkinter import Tk, filedialog
from pptx import Presentation
from docx import Document
from tqdm import tqdm
import jieba,json,os,joblib,PyPDF2

#读取pptx文件
def pptx(name):
    prs = Presentation(str(name))
    text_runs = []
    for slide in prs.slides:
        for shape in slide.shapes:
            if not shape.has_text_frame:
                continue
            for paragraph in shape.text_frame.paragraphs:
                for run in paragraph.runs:
                    text_runs.append(run.text)
    return text_runs

#读取docx文件
def docx(name):
    document = Document(str(name))
    ps = [ paragraph.text for paragraph in document.paragraphs]
    return ps

#读取pdf文件
def pdf(name):
    mypdf = open(name, mode='rb')
    pdf_document = PyPDF2.PdfReader(mypdf)
    an = []
    for i in range(len(pdf_document.pages)):
        t = pdf_document.pages[i].extract_text().split(' ')
        if len(t)>1:
            an += t
    return an

#对读取的数据进行预处理
def structure(lst,md=2):
    def chdl(r):
        r = jieba.lcut(r)
        i = 0
        while i<len(r):
            if len(r[i])<2:
                r.pop(i)
                continue
            i += 1
        return r
    
    ls = []
    for i in lst:
        if len(i)<2:
            continue
        cnt = [0]*2
        x = ['']*2
        i = i.lower()
        for s in i:
            if '\u4e00' <= s <= '\u9fff':
                cnt[0] += 1
                x[0] += s
            elif 'a' <= s <= 'z':
                cnt[1] += 1
                x[1] += s
            else:
                if len(x[0])>1: ls += chdl(x[0])
                if len(x[1])>2: ls.append(x[1])
                x = ['']*2
        if len(x[0])>1: ls += chdl(x[0])
        if len(x[1])>2: ls.append(x[1])
    if md==2:
        return ' '.join(ls)
    dic = {}
    for i in ls:
        if i in dic:
            dic[i] += 1
        else:
            dic[i] = 1
    return dic

#通过文件名调用对应函数
def getcont(fn,md=2):
    typ = fn.split('.')[-1]
    if typ=='pptx':
        a = pptx(fn)
    elif typ=='docx':
        a = docx(fn)
    elif typ=='pdf':
        a = pdf(fn)
    else:
        return False
    return structure(a,md)

'''
#合并或创建DataFrame词频数据（暂时无用）
def merge(df,dic,col):
    if df is None:
        df = pd.DataFrame(list(dic),columns=['word'])
    if col not in df.columns:
        df[col] = 0
    for i in dic:
        if (df.word==i).any():
            df.loc[len(df)] = [i,0]
        df.loc[df.word==i,col] += dic[i]
    return df

#手动更改分类类型（暂时无用）
def setp(r):
    while True:
        print()
        for i in range(len(r)):
            print(f'{i+1}.',r[i],end=' ')
        print('\n1.add 2.del 3.quit')
        try:
            ip = input()
            if ip=='1':
                print('直接回车退出')
                while len(r)==0 or r[-1]!='':
                    r.append(input())
                r.pop()
            elif ip=='2':
                r.pop(int(input())-1)
            elif ip=='3':
                break
            else:
                print('error')
        except:
            print('error')
    return r
'''

#模型的训练与保存
def model_sv(rec):
    doc = []
    labels = []
    print('开始预处理数据')
    for i in tqdm(rec):
        for j in tqdm(rec[i]):
            try:
                doc.append(getcont(j))
                labels.append(i)
            except:
                tqdm.write(f'添加数据失败：{j}')
    print('开始训练模型')
    vectorizer = TfidfVectorizer()
    X = vectorizer.fit_transform(doc)
    X_train, X_test, y_train, y_test = train_test_split(X, labels, test_size=0.3, random_state=42)
    clf = MultinomialNB()
    clf.fit(X_train, y_train)
    y_pred = clf.predict(X_test)
    accuracy = accuracy_score(y_test, y_pred)
    print(f"模型准确率: {accuracy}")
    del clf
    clf = MultinomialNB()
    clf.fit(X, labels)
    while True:
        n = input('请命名此模型：')
        t = os.path.join(tmp['modph'],n)
        vt,cf = os.path.join(t,'vectorizer.pkl'),os.path.join(t,'classifier.pkl')
        if os.path.exists(vt) or os.path.exists(cf):
            print('已有此名字模型，请重新输入\n')
            continue
        if not os.path.exists(t):
            os.makedirs(t)
        try:
            joblib.dump(vectorizer, vt)
            joblib.dump(clf, cf)
            break
        except:
            print('命名不符合规范\n')
    print('模型保存成功')

#模型的使用
def model_use():
    f = []
    lis = []
    for i in tqdm(os.listdir(config['collect'])):
        if not os.path.isdir(i) and i.split('.')[-1] in ['pptx','docx','pdf']:
            try:
                t = getcont(os.path.join(config['collect'],i))
                if t!=False:
                    f.append(i)
                    lis.append(t)
            except:
                pass
    vectorizer = joblib.load(os.path.join(tmp['modph'],config['nmod'],'vectorizer.pkl'))
    clf = joblib.load(os.path.join(tmp['modph'],config['nmod'],'classifier.pkl'))
    prediction = clf.predict(vectorizer.transform(lis))
    for i in range(len(f)):
        t = os.path.join(config['save'],prediction[i])
        if not os.path.exists(t):
            os.makedirs(t)
        if config['sepmd']==1:
            os.rename(os.path.join(config['collect'],f[i]),os.path.join(t,f[i]))
        elif config['sepmd']==2:
            copy_file(os.path.join(config['collect'],f[i]),os.path.join(t,f[i]))

#复制文件
def copy_file(src, dst):
    with open(src, 'rb') as fsrc:
        with open(dst, 'wb') as fdst:
            while True:
                buf = fsrc.read(1024)
                if not buf:
                    break
                fdst.write(buf)

#将文件夹路径转换为类别与文件路径的对应数据
def fold2file(n):
    lis = {}
    for i in os.listdir(n):
        i2 = os.path.join(n,i)
        if not os.path.isdir(i2):
            continue
        lis[i] = []
        for dirp,dirn,filen in os.walk(i2):
            for j in filen:
                k = os.path.join(dirp,j)
                if k.split('.')[-1] in ['pptx','docx','pdf']:
                    lis[i].append(k)
    return lis

#使用可视化文件夹选择
def select_directory(tip):
    root = Tk()
    root.withdraw()
    return filedialog.askdirectory(title=tip)

#读取配置文件
def config_rd():
    global config
    if not os.path.exists('data'):
        os.makedirs('data')
    cfgph = os.path.join('data','config.json')
    if os.path.exists(cfgph):
        with open(cfgph, 'r') as file:
            config = json.load(file)

#保存配置文件
def config_sv():
    global config
    with open(os.path.join('data','config.json'), 'w') as file:
        json.dump(config,file)

#修改配置
def change_config(r):
    global config
    if r=='collect':
        tp = select_directory('请选择文档“收集”文件夹，取消则默认为桌面')
        if tp=='':
            tp = os.path.join(os.path.expanduser('~'), 'Desktop')
        config['collect'] = tp
    elif r=='save':
        tp = select_directory('请选择文档“存放”文件夹，取消则默认为“收集”文件夹内的“已分类文档”文件夹')
        if tp=='':
            tp = os.path.join(config['collect'], '已分类文档')
            if not os.path.exists(tp):
                os.makedirs(tp)
        config['save'] = tp
    elif r=='nmod':
        nm = []
        for i in os.listdir(tmp['modph']):
            if os.path.isdir(os.path.join(tmp['modph'],i)):
                nm.append(i)
        while True:
            for i in range(len(nm)):
                print(f'{i+1}.',nm[i])
            try:
                ip = int(input('请选择模型：'))
                t = os.path.join(tmp['modph'],nm[ip-1])
                if os.path.exists(os.path.join(t,'vectorizer.pkl')) and os.path.exists(os.path.join(t,'classifier.pkl')):
                    config['nmod'] = nm[ip-1]
                    break
                print('没有此模型')
            except:
                print('error')
    elif r=='sepmd':
        while True:
            try:
                print('\n请选择处理方式：1.移动 2.复制')
                ip = int(input('请输入：'))
                if ip not in [1,2]:
                    print('error')
                    continue
                config['sepmd'] = ip
                break
            except:
                print('error')
    config_sv()

#程序初始化
def initialize():
    config_rd()
    tmp['modph'] = os.path.join('data','model')
    if not os.path.exists(tmp['modph']):
        os.makedirs(tmp['modph'])
    if config['collect']=='' or not os.path.exists(config['collect']):
        change_config('collect')
    if config['save']=='' or not os.path.exists(config['save']):
        change_config('save')
    t = os.path.join(tmp['modph'],config['nmod'])
    if not (os.path.exists(os.path.join(t,'vectorizer.pkl')) and os.path.exists(os.path.join(t,'classifier.pkl'))):
        config['nmod'] = ''
        config_sv()

#外部数据与必要设置
ver = 'v1.0.0 beta'
jieba.set_dictionary(os.path.join('data','dict.txt'))
config = {'nmod':'default','collect':'','save':'','sepmd':1}
tmp = {}
initialize()

#主程序
if __name__=='__main__':
    print('本程序完全免费\nGithub项目地址：https://github.com/ying-ck/Document-classifier\n作者：Yck')
    print('当前版本:',ver)
    while True:
        print('\n选择功能：1.训练模型 2.进行分类 3.设置 4.退出')
        ip = input('请输入：')
        if ip=='1':
            t = select_directory('请选择文件夹')
            if t!='':
                model_sv(fold2file(t))
        elif ip=='2':
            t = os.path.join(tmp['modph'],config['nmod'])
            if config['nmod']=='' or not (os.path.exists(os.path.join(t,'vectorizer.pkl')) and os.path.exists(os.path.join(t,'classifier.pkl'))):
                change_config('nmod')
            print('开始分类')
            model_use()
            print('分类结束')
        elif ip=='3':
            print('\n选择：1.文档收集位置 2.文档存放位置 3.选择模型 4.文档分类处理方式')
            ip = input()
            if ip=='1':
                change_config('collect')
            elif ip=='2':
                change_config('save')
            elif ip=='3':
                change_config('nmod')
            elif ip=='4':
                change_config('sepmd')
        elif ip=='4':
            break

