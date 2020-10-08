import time
import sys
import shutil
import numpy as np
import os
import warnings
warnings.filterwarnings(action='ignore',category=UserWarning,module='gensim')
from PyQt5.QtWidgets import QFileDialog
from PyQt5.uic import loadUi
from PyQt5.QtWidgets import QMainWindow, QApplication
import jieba.posseg as pseg
import codecs#文件读取时，编码转化
from gensim import corpora, models, similarities#运用Gensim建立词典，生成BOW语料
#运行tfidf模型计算词权重，采用LsiModel进行降维，最后运用Gensim提供的MatrixSimilarity类来计算两文档的相似性【基于余弦的距离的计算】
import docx2txt#读取docx文档
import xlwt
import easygui
from QCandyUi.CandyWindow import colorful
@colorful('blueGreen')
class MainWindow(QMainWindow):
    def __init__(self):
        QMainWindow.__init__(self)
        loadUi("check.ui", self)
        global FilesName
        global Directory
        global Name
        global threshold
        
        #botton 3
        FilesName =self.pushButton_3.clicked.connect(self._openFile_)
        #botton 2
        Directory=str(self.pushButton_2.clicked.connect(self._openDir_))
        #botton 1


        self.pushButton.clicked.connect(lambda:self._check_(FilesName,Directory))

        
    def _split_(self,file_path):
        (filepath, tempfilename) = os.path.split(file_path)
        (filename, extension) = os.path.splitext(tempfilename)
        return filename

    def _check_(self,FilesName,Directory):
        #报表名前后无空格
        global Name
        Name=str(self.lineEdit_2.text()).strip()+'.xls'
        global threshold
        threshold=70 #default
        
        s=self.lineEdit.text().strip()
        if s!='':
            threshold=int(s)
            
        
       
        
        style = xlwt.easyxf('pattern: pattern solid, fore_colour red')
        stop_words = 'cn_stopwords.txt'#停用词路径
        stopwords = codecs.open(stop_words,'r',encoding='utf8').readlines()#读取停用词
        stopwords = [ w.strip() for w in stopwords ] # 去除首尾空格
        stop_flag = ['x', 'c', 'u','d', 'p', 't', 'uj', 'm', 'f', 'r']# 停用词性

        #对一篇文章分词、去停用词
        def tokenization(file):
            
            result = []
            text = file
            words = pseg.cut(text)#分词
            #去除 flag在stop_flag，word 在stopword，所对应的词语（即除二者意外其他加入了result中）
            for word, flag in words:
                if flag not in stop_flag and word not in stopwords:
                    result.append(word)
            return result

       
        
            #对目录下的所有文本进行预处理，构建字典
        start =time.clock()
        filenames=[]
        for i in FilesName:
            filenames.append(docx2txt.process(i))

        

        corpus = []
        for each in filenames:
            corpus.append(tokenization(each))


        dictionary = corpora.Dictionary(corpus)#为每个出现在词库中的词语分配一个独一无二的整数ID，这个操作收集单词计数和其他统计信息


        #建立词袋模型
        #  生成词向量
        doc_vectors = [dictionary.doc2bow(text) for text in corpus] # 语料库

        # 生成TF-IDF模型
        tfidf = models.TfidfModel(doc_vectors)# TF-IDF模型对语料库建模
        tfidf_vectors = tfidf[doc_vectors] # 每个词的TF-IDF值


        #把稀疏的高级矩阵变成一个计算起来会比较轻松的小矩阵，也把一些没有用的燥音给过滤掉了，这个模型可以被后来的语料查询与分类所调用
        lsi = models.LsiModel(tfidf_vectors, id2word=dictionary, num_topics=300)
        #num_topics 主题数
        #将Tf-Idf语料转化为一个潜在2-D空间（2-D是因为我们设置了num_topics=2）

        lsi.print_topics(300)# 建立的两个主题模型内容

        # 将文章投影到主题空间中
        lsi_vector = lsi[tfidf_vectors]
        count1=0
        count=1
        workbook=xlwt.Workbook(encoding="utf-8")
        worksheet=workbook.add_sheet(Name)
        worksheet.write(0,0,'查找文档名')
        worksheet.write(0,1,'对比文档名')
        worksheet.write(0,2,'重复率')
        for i in FilesName:
            
            text = docx2txt.process(i)#打开文档
            query = tokenization(text)#分词，去词
            query_bow = dictionary.doc2bow(query)#语义库

            query_lsi = lsi[query_bow]


            #相似矩阵计算相似度
            index = similarities.MatrixSimilarity(lsi_vector)#和第个文档相似
            sims = index[query_lsi]

            sort_sims = sorted(enumerate(sims), key=lambda item: -item[1])
            
            worksheet.write(count,0,self._split_(i))
            worksheet.write(count,1,self._split_(FilesName[sort_sims[1][0]]))
            if np.float32(sort_sims[1][1]).item()>(float(threshold)*0.01):
                 worksheet.write(count,2,np.float32(sort_sims[1][1]).item(),style)
                 count1=count1+1
            else:
                 worksheet.write(count,2,np.float32(sort_sims[1][1]).item())
           
            
            count=count+1
        
        workbook.save(Name)
        
        aa=os.getcwd()
        file_path=os.path.join(aa,Name)
        shutil.move(file_path,Directory)


       
        end = time.clock()
        print('Running time: %s Seconds'%(end-start))
        easygui.msgbox("完成啦！", title="文本查重",ok_button="知道啦")
        
      
    #button3 选择文件，返回list类型的文件路径
        
    def _openFile_(self):
        global FilesName
        
        FilesName,filetype = QFileDialog.getOpenFileNames(None, "请选择要添加的文件", "C:\\Users\\14866\\Desktop\\NLP\\Untitled Folder", "Text Files (*.docx);;Text Files (*.doc);;All Files (*)")
        return FilesName

    def _openDir_(self):
        global Directory
        Directory = QFileDialog.getExistingDirectory(self, "选择文件夹", "/")
        return Directory


    



app = QApplication(sys.argv)
window = MainWindow()
window.show()
sys.exit(app.exec_())
