import os
import re
import sys
import docx 
import glob
import nltk
import spacy
import PyPDF4
import warnings
import qdarkstyle
import pdfplumber
import pandas as pd
import PyQt5.QtGui as qtg
import PyQt5.QtCore as qtc
import PyQt5.QtWidgets as qtw

from fuzzywuzzy import fuzz
from spacy.tokens import Span
from fuzzywuzzy import process
from nltk.corpus import stopwords
from urlextract import URLExtract
from pyresparser import ResumeParser
from spacy.pipeline import EntityRuler
from nltk.stem import WordNetLemmatizer

nlp=spacy.load('en_core_web_md')
warnings.filterwarnings('ignore')
nltk.download(['stopwords','wordnet'])
skill_pattern_path = "jz_skill_patterns.jsonl"
ruler=EntityRuler(nlp)
ruler.from_disk(skill_pattern_path)
ruler = nlp.add_pipe(ruler)
class resumes:
    def __init__(self,main_dir,job_description):
        try:
            os.mkdir('raw')
        except:
            pass
        try:
            os.mkdir('formated_info')
        except:
            pass

        

        self.main_dir=main_dir
        df=pd.read_csv('world-universities.csv')
        df.columns=['_',"University","email"]
        self.universities=df["University"]
        with open('abrivs.txt','r') as f:
            abrivs=f.read()
        self.abrivs=abrivs.replace('\n',' ')

        self.docx_files=glob.glob(self.main_dir+'/*.docx')
        self.pdf_files=glob.glob(self.main_dir+'/*.pdf')
        self.extractor = URLExtract()
        self.texts=[]
        self.infos=[]
        self.names=[]
        self.job_description=job_description
        self.skills_per=dict()
        self.skills_required=self.get_required_skills()
    def get_required_skills(self):
        if '.pdf' in self.job_description:
            text=self.getText_pdf(self.job_description)
        elif '.docx' in self.job_description:
            text=self.getText_docx(self.job_description)
        elif '.txt' in self.job_description:
            
            try:
                with open(self.job_description,'r') as f:
                    text=f.read()
            except:
                try:
                    with open(self.job_description,'r',encoding='utf-8') as f:
                        text=f.read()
                except:
                    pass
        skills,education=self.get_skills_education(text)  
        return ' '.join(skills).lower()
    def get_matches(self):
        skills_required=nlp(self.skills_required)
        ratios=dict()
        for i in self.skills_per:
            person_skills=self.skills_per[i]
            person_skills=nlp(person_skills)
            match=person_skills.similarity(skills_required)
            match=round(match*100,2)
            ratios[i]=match
        return ratios
    
    
    
    def extract_data(self):
        for file in self.pdf_files:
            info=[]
            name=file.split('\\')[-1].split('.')[0]
            self.names.append(name)
            data = ResumeParser(file).get_extracted_data()
            skills_1=data['skills']
            text = self.getText_pdf(file)
            skills,education=self.get_skills_education(text)
            skills=list(set(skills+skills_1))
            emails = re.findall(r"[a-z0-9\.\-+_]+@[a-z0-9\.\-+_]+\.[a-z]+", text)
            emails.append(data['email'])
            emails=list(set(emails))
            links = self.extractor.find_urls(text)
            links = self.clean_list(links,self.abrivs)
            self.texts.append(text)
            info.append("filename: "+file)
            info.append("Name: "+name)
            self.skills_per[name]=' '.join(skills)
            try:
                info.append("Skills: "+' , '.join(skills))
            except:
                info.append("Skills: Not Found")
            try:
                info.append("Emails: "+' , '.join(emails))
            except:
                info.append("Emails: Not Found")
            try:
                info.append("Links: "+' , '.join(links))
            except:
                info.append("Links: Not Found")
            self.infos.append(info)

        for file in self.docx_files:
            info=[]
            name=file.split('\\')[-1].split('.')[0]
            self.names.append(name)
            info.append("filename: "+file)
            info.append("Name: "+name)
            data = ResumeParser(file).get_extracted_data()
            skills_1=data['skills']
            skills_1=' '.join(skills_1)
            skills_1=skills_1.lower()
            skills_1=skills_1.split()
            text = self.getText_docx(file)
            skills,education=self.get_skills_education(text)
            skills=list(set(skills+skills_1))
            self.skills_per[name]=' '.join(skills)
            emails = re.findall(r"[a-z0-9\.\-+_]+@[a-z0-9\.\-+_]+\.[a-z]+", text)
            emails.append(data['email'])
            emails=list(set(emails))
            links = self.extractor.find_urls(text)
            links = self.clean_list(links,self.abrivs)
            self.texts.append(text)
            try:
                info.append("Skills: "+' , '.join(skills))
            except:
                info.append("Skills: Not Found")
            try:
                info.append("Emails: "+' , '.join(emails))
            except:
                info.append("Emails: Not Found")
            try:
                info.append("Links: "+' , '.join(links))
            except:
                info.append("Links: Not Found")
            self.infos.append(info)

        for i,text in enumerate(self.texts):
            with open(f'raw/{self.names[i]}.txt','w',encoding='utf-8') as f:
                f.write(text)
        for i,info in enumerate(self.infos):
            with open(f'formated_info/{self.names[i]}.txt','w') as f:
                f.write('\n\n'.join(info))

    def getText_docx(self,filename):
        doc = docx.Document(filename)
        fullText = []
        for para in doc.paragraphs:
            fullText.append(para.text)
        return ' '.join(fullText).replace('\n',' ')

    def getText_pdf(self,filename):
        fullText=[]
        with pdfplumber.open(filename) as pdf:
            pages = pdf.pages
            for page in pages:
                fullText.append(page.extract_text())
        return '\n'.join(fullText).replace('\n',' ')
    def clean_text(self,text):
        text=re.sub(
            '(@[A-Za-z0-9]+)|([^0-9A-Za-z \t])|(\w+:\/\/\S+)|^rt|http.+?"',
            " ",text)
        #text=text.lower()
        text=text.split()
        lm = WordNetLemmatizer()
        text = [
            lm.lemmatize(word)
            #word
            for word in text
            if not word in set(stopwords.words("english"))
        ]
        text = " ".join(text)
        return text
    def get_skills_education(self,text):
        skills=[]
        education=[]
        for i in self.universities:
            if i.lower() in text.lower() :
                education.append(i)
        textt=self.clean_text(text)
        textt=nlp(textt)
        for ent in textt.ents:
            if ent.label_ == "SKILL":
                skills.append(ent.text.lower())
        skills=list(set(skills))
        education=list(set(education))
        return skills,education
    def clean_list(self,my_list,comp):
        MyList=[]
        for i in my_list:
            if i.lower() in comp.lower():
                pass
            else:
                MyList.append(i)
        return MyList
    
class MainWindow(qtw.QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("CV shortlisting and ranking")
        self.resize(800,800)
        
        self.select_file=qtw.QPushButton("Select Job description file")
        self.select_file.setStyleSheet(self.btn_style())
        self.select_file.clicked.connect(self.file_dialog)
        
        self.select_folder=qtw.QPushButton("Select Resumes directory")
        self.select_folder.setStyleSheet(self.btn_style())
        self.select_folder=qtw.QPushButton("Select Resumes directory")
        self.select_folder.setStyleSheet(self.btn_style())
        self.select_folder.clicked.connect(self.get_dir)
        
        self.load_res=qtw.QPushButton("load Resumes")
        self.load_res.setStyleSheet(self.btn_style())
        self.load_res.clicked.connect(self.get_matches)
        
        self.show_res=qtw.QPushButton("Show Results")
        self.show_res.setStyleSheet(self.btn_style())
        self.show_res.clicked.connect(self.visit)
        
        self.back=qtw.QPushButton("Back to Home page")
        self.back.setStyleSheet(self.btn_style())
        self.back.clicked.connect(self.back_home)

        ###################################################
        self.skills=qtw.QLineEdit()
        self.skills_label=qtw.QLabel('Enter skills you want:')
        self.skills_label.setStyleSheet(self.label_style())
        
        self.skills.setStyleSheet(self.line_edit_style())
        self.skills.editingFinished.connect(self.add_word)
        
        
        self.clear=qtw.QPushButton("Clear")
        self.clear.clicked.connect(self.clear_list)
        self.clear.setStyleSheet(self.btn_style())
        
        self.remove_keyword=qtw.QPushButton('remove')
        self.remove_keyword.clicked.connect(self.remove_word)
        self.remove_keyword.setStyleSheet(self.btn_style())
        
        self.skills_list=qtw.QListWidget()
        
        self.add_keyword=qtw.QPushButton('Add skill')
        self.add_keyword.setStyleSheet(self.btn_style())
        self.add_keyword.clicked.connect(self.add_word)
        
        self.save_list=qtw.QPushButton('save skills')
        self.save_list.setStyleSheet(self.btn_style())
        self.save_list.clicked.connect(self.save_skills)
        
        self.tabel=qtw.QTableWidget(4,2)
        self.tabel.setColumnWidth(0,400)
        self.tabel.setColumnWidth(1,100)
        
        
        self.setup_layout()
    def save_skills(self):
        l=[]
        for i in range(self.skills_list.count()):
            l.append(self.skills_list.item(i).text())
        l=' '.join(l)
        if len(l)>0:
            with open('skills.txt','w') as f:
                f.write(l)
                
    def get_matches(self):
        
        my_resume=resumes(self.resumes_dir,self.job_disc_dir)
        my_resume.extract_data()
        matches=my_resume.get_matches()
        matches=dict(sorted(matches.items(), key=lambda item: item[1]))
        self.tabel.setRowCount(len(matches))
        for ind,i in enumerate(matches):
            self.tabel.setItem(ind,0,qtw.QTableWidgetItem(str(i)))
            self.tabel.setItem(ind,1,qtw.QTableWidgetItem(str(matches[i])))
        
    def visit(self):
        self.stckd.setCurrentIndex(1)
    def back_home(self):
        self.stckd.setCurrentIndex(0)
    def clear_list(self):
        self.skills_list.clear()
    def label_style(self):
        return '''QLabel{
        color:white;
        font:bold 12px;
        min-width=5em;
        }
        '''
    def remove_word(self):
        self.items=self.skills_list.selectedItems()
        for item in self.items:
            QLWI=self.skills_list.takeItem(self.skills_list.row(item))
    def add_word(self):
        self.skills_list.addItem(qtw.QListWidgetItem(self.skills.text(),self.skills_list))
        self.skills.setText('')
    def line_edit_style(self):
        return'''
        color:white;
        min-height:2em;
        border-radius:5px;
        font:bold 14px;
        min-width:25em
        '''
    def btn_style(self):
        return """
               QPushButton{
                    background-color:blue;
                    color:white;
                    border-radius:10px;
                    min-width:14em;
                    min-height:2em;
                    font:bold 14px;
                }

                QPushButton::hover{
                    background-color:rgb(50,250,250);
                    color:white;
                    border-radius:10px;
                    min-width:14em;
                    min-height:3em;
                    font:bold 14px;
                }
                """
    def file_dialog(self):
        res=qtw.QFileDialog.getOpenFileName(self,'Open File','C:\\',"Image files (*.pdf *.docx *.txt)")
        self.job_disc_dir=res[0]
    def get_dir(self):
        file = str(qtw.QFileDialog.getExistingDirectory(self, "Select Directory"))
        self.resumes_dir=file

    def setup_layout(self):
        
        self.stckd=qtw.QStackedLayout()
        self.main_w=qtw.QWidget()
        self.home_layout=qtw.QHBoxLayout()
        self.left_ly=qtw.QVBoxLayout()
        self.left_ly.addWidget(self.select_folder)
        self.left_ly.addWidget(self.select_file)
        self.left_ly.addWidget(self.load_res)
        self.left_ly.addWidget(self.show_res)
        self.home_layout.addLayout(self.left_ly)
        self.home_layout.addStretch()
        
        self.right_layout=qtw.QVBoxLayout()
        self.right_layout.addWidget(self.skills_list)
        self.right_layout.addWidget(self.skills_label)
        self.right_layout.addWidget(self.skills)
        self.buttons_layout=qtw.QHBoxLayout()
        self.buttons_layout.addWidget(self.add_keyword)
        self.buttons_layout.addWidget(self.remove_keyword)
        self.buttons_layout.addWidget(self.clear)
        self.right_layout.addLayout(self.buttons_layout)
        self.right_layout.addWidget(self.save_list)
        self.home_layout.addLayout(self.right_layout)
        self.main_w.setLayout(self.home_layout)
        
        
        
        self.tabel_ly_w=qtw.QWidget()
        self.tabel_ly=qtw.QVBoxLayout()
        self.tabel_ly.addWidget(self.back)
        self.tabel_ly.addWidget(self.tabel)
        self.tabel_ly.addStretch()
        self.tabel_ly_w.setLayout(self.tabel_ly)

        
        
        
        self.stckd.addWidget(self.main_w)
        self.stckd.addWidget(self.tabel_ly_w)
        self.setLayout(self.stckd)
        self.stckd.setCurrentIndex(0)
        
        
        self.setLayout(self.home_layout)
    
        self.show()
if __name__=='__main__':
    app=qtw.QApplication([])
    app.setStyleSheet(qdarkstyle.load_stylesheet_pyqt5())
    mw=MainWindow()
    app.exec_()
   