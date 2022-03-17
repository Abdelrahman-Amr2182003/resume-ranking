import os#importing the os library to be able to creat folder
import re#import tre to recognize patterns
import docx#to read text docx files
import glob# to get all filenames inside the directory
import nltk#to lemmatize words and filter the stop words
import spacy# used to get similarity tokenize and preprocess the text
import warnings# to be able to ignore warning
import qdarkstyle# to have the application be in darkstyle
import pdfplumber# read text from pdf
import pandas as pd# read csv files
import PyQt5.QtGui as qtg#to be able to assign icons to buttons and labels and other important things
import PyQt5.QtWidgets as qtw# import all the Qt widgets

from nltk.corpus import stopwords# to get a list of all stopwords in the nltk library
from urlextract import URLExtract# extract urls from text
from pyresparser import ResumeParser# parse through the resume to extract important data to use alongside our alogrithem so that we get both outputs of both alogriths withouth reptitions
from spacy.pipeline import EntityRuler# to add a new entity "Skills"
from nltk.stem import WordNetLemmatizer# to lemmatize the word (return to its origin form)

nlp=spacy.load('en_core_web_md')#load the middle english library from spacy
warnings.filterwarnings('ignore')#ignore all warnings
nltk.download(['stopwords','wordnet'])#download stopwords
skill_pattern_path = "jz_skill_patterns.jsonl"#path to the skill patern data set that contains most of the skills to most jobs
ruler=EntityRuler(nlp)#define a new EntityRuler
ruler.from_disk(skill_pattern_path)#load dataset to ruler
ruler = nlp.add_pipe(ruler)# add skills as a new entity
class resumes:# creat a class caleed resumes
    def __init__(self,main_dir,job_description):# the constructor of the class takes two parameters 1-directory to resumes folder ,path to job description file

        try:# create new folders if not already there if it is then go on 
            os.mkdir('raw')
        except:
            pass
        try:
            os.mkdir('formated_info')
        except:
            pass

        

        self.main_dir=main_dir#define self.main as the resumes directory
        df=pd.read_csv('world-universities.csv')# read a csv file containg all universites in the world
        df.columns=['_',"University","email"]#set column names
        self.universities=df["University"]# get list of all universities
        
        with open('abrivs.txt','r') as f:# read text file containing all abrivitions to execlude from bein urls (because the contain ".")
            abrivs=f.read()
        self.abrivs=abrivs.replace('\n',' ')# replace the new line charcter with a space

        self.docx_files=glob.glob(self.main_dir+'/*.docx')# get all docx filpaths
        self.pdf_files=glob.glob(self.main_dir+'/*.pdf')# get all pdf filepaths
        self.extractor = URLExtract()#make an object of the URLExtract class
        self.texts=[]#empty list to store texts
        self.infos=[]#em[ty list to store informations
        self.names=[]# emty list to store the names of people
        self.job_description=job_description#define self.job_description as the path to the job_description file
        self.skills_per=dict()#dictionary to store skills each person have
        self.skills_required=self.get_required_skills()#get all the skills in the job description
    def get_required_skills(self):
        if '.pdf' in self.job_description:# if job description is a pdf fileextract the text using the getText_pdf function
            text=self.getText_pdf(self.job_description)
        elif '.docx' in self.job_description:# if job description is a pdf fileextract the text using the getText_docx function
            text=self.getText_docx(self.job_description)
        elif '.txt' in self.job_description:# if job description is a pdf fileextract the text using the getText_pdf functionopen(filename) function
            try:
                with open(self.job_description,'r') as f:#try reading the file with default encoding
                    text=f.read()# get the text
            except:
                try:
                    with open(self.job_description,'r',encoding='utf-8') as f:# if it got in error the read the file with the utf-8 encoding
                        text=f.read()# get the text
                except:#if not then the file is not corupted so dont do anything
                    pass
        skills,_=self.get_skills_education(text)# extract requiredd text from the job description
        return ' '.join(skills).lower()# return a space seprated string of all the skills 
    def get_matches(self):# define the get_matches function
        skills_required=nlp(self.skills_required)#transofrom the string of the skills required to a spacy token object
        ratios=dict()#define a dictionary that stores the match of each person
        for i in self.skills_per:#iterate through the keys(names) of the skills_per dictionary 
            person_skills=self.skills_per[i]# get skills assigned to that person
            person_skills=nlp(person_skills)#transofrom the string of the skills to a spacy token object
            match=person_skills.similarity(skills_required)# get cosine similarity between person skills and skills_required
            match=round(match*100,2)# get it to be in percentile form and round only the first two digits after the decimal point
            ratios[i]=match#make that the match of the person equal to the result of the cosine similarity(match variable)
        return ratios# return ratios dictionary
    
    
    
    def extract_data(self):#define the extract data function
        
        for file in self.pdf_files:#iterate through all pdf files
            info=[]# empty list to sva einformation regarding each file
            name=file.split('\\')[-1].split('.')[0]# get the name of the person by removing the directory path and the file format leaving only the filename 
            #which is supposdly the person name
            self.names.append(name)# append name to the names list
            data = ResumeParser(file).get_extracted_data()# parse the resume using the ResumeParser class the algorithm we are using is better however it may 
            #extract info that wasnt extracted by our alogrithm so we add both to each others
            skills_1=data['skills']# get skills extracted from the ResumeParser class
            text = self.getText_pdf(file)#extract the text using the getText_pdf function 
            skills,education=self.get_skills_education(text)# extract the skills and the education using the get_skills_education function
            skills=list(set(skills+skills_1))# add skills extract from both alogrithms together and remove duplicates
            emails = re.findall(r"[a-z0-9\.\-+_]+@[a-z0-9\.\-+_]+\.[a-z]+", text)# extract emails by trying to find this pattern
            emails.append(data['email'])#appen email found from ResumeParser to emails
            emails=list(set(emails))#remove duplicates
            links = self.extractor.find_urls(text)# get urls
            links = self.clean_list(links,self.abrivs)# remove abriviations from the list of urls
            self.texts.append(text)# append text to texts
            info.append("filename: "+file)# append filename to info
            info.append("Name: "+name)# append name to info
            self.skills_per[name]=' '.join(skills)
            try:
                info.append("Skills: "+' , '.join(skills))# append skills
            except:
                info.append("Skills: Not Found")
            try:
                info.append("Emails: "+' , '.join(emails))#append emails
            except:
                info.append("Emails: Not Found")
            try:
                info.append("Links: "+' , '.join(links))#append links
            except:
                info.append("Links: Not Found")
            self.infos.append(info)

        for file in self.docx_files:# same for pdf but with docx file
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
            text = self.getText_docx(file)# the only diffrence is we extract ext using the getText_docx function 
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

        for i,text in enumerate(self.texts):# creat new text file containg raw text
            with open(f'raw/{self.names[i]}.txt','w',encoding='utf-8') as f:# save file with encoding='utf-8'
                f.write(text)#load text into the file
        for i,info in enumerate(self.infos):## creat new text file containg processed text
            with open(f'formated_info/{self.names[i]}.txt','w') as f:# save file
                f.write('\n\n'.join(info))#load text into files and seprate each inofrmation by two lines

    def getText_docx(self,filename):# define the getText_docx function
        doc = docx.Document(filename)# use the docx library to load the word document here
        fullText = []#empty list to stroe text
        for para in doc.paragraphs:# iterate through all the paragraphs
            fullText.append(para.text)# append text to fullText 
        return ' '.join(fullText).replace('\n',' ')#join elemnts of text with ' ' and replace new line charcter with a ' ' and return the result

    def getText_pdf(self,filename):#define getText_pdf function
        fullText=[]# empty text 
        with pdfplumber.open(filename) as pdf:# open the pdf file using the pdfplumber library
            pages = pdf.pages#get list of all pages inside the pdf
            for page in pages:#iterate through all pages
                fullText.append(page.extract_text())#append page text to fullText
        return '\n'.join(fullText).replace('\n',' ')#join elemnts of text with ' ' and replace new line charcter with a ' ' and return the result
    
    def clean_text(self,text):# define the clean_text
        text=re.sub(
            '(@[A-Za-z0-9]+)|([^0-9A-Za-z \t])|(\w+:\/\/\S+)|^rt|http.+?"',
            " ",text)# remove garbage characters from list 
        #text=text.lower()
        text=text.split()#split the text into a list
        lm = WordNetLemmatizer()#define an object of WordNetLemmatizer class
        text = [
            lm.lemmatize(word)#lematization
            #word
            for word in text#iterate through the whole text
            if not word in set(stopwords.words("english"))#return the word only if not from the stopwords
        ]#lematize words and filter stopwords
        text = " ".join(text)# join list inot string 
        return text#return clean text
    def get_skills_education(self,text):# define the get_skills_education function
        skills=[]#empty list to store skills
        education=[]#empty list to stor universties the applicant went to
        for i in self.universities:# iterate through all univesties 
            if i.lower() in text.lower() :#if one of them in the text
                education.append(i)#append it to the list
        textt=self.clean_text(text)#clean the text using the clean_text function
        textt=nlp(textt)#transofrm to spacy token object
        for ent in textt.ents:# iterate through entities
            if ent.label_ == "SKILL":#if entity is a skill
                skills.append(ent.text.lower())#append to list of skills
        skills=list(set(skills))#remove duplicates
        education=list(set(education))#remove duplicates
        return skills,education#return skills and education
    def clean_list(self,my_list,comp):# define the clean_list function which filter words in comp from my_list
        MyList=[]# empty list to store good words
        for i in my_list:
            if i.lower() in comp.lower():#if word in comp_string then its not neeeded
                pass#do nothing
            else:# if not
                MyList.append(i)#append to MyList
        return MyList# return MyList
    
class MainWindow(qtw.QWidget):# define the MainWindow class for the user interface and saying that it inherits from the QWidget class
    def __init__(self):# init function ofr initilizations
        super().__init__()#inherit the QWidget class functions and attributes 
        self.setWindowTitle("CV shortlisting and ranking")# set the title
        self.resize(800,800)# window size
        
        self.select_file=qtw.QPushButton("Select Job description file")# define a new button with a label . the button loads the job_description file
        self.select_file.setStyleSheet(self.btn_style())# set the style to the return of the btn_style funtion (the css style sheet)
        self.select_file.clicked.connect(self.file_dialog)# connect button to the file_dialog function
        
        
        self.select_folder=qtw.QPushButton("Select Resumes directory")# define a new button with a label. the button reads the folder where the resumes are
        self.select_folder.setStyleSheet(self.btn_style())# set the style to the return of the btn_style funtion (the css style sheet)
        self.select_folder.clicked.connect(self.get_dir)# connect button to the get_dir function
        
        self.load_res=qtw.QPushButton("Proces resumes")# define a new button with a label. the button process adn extract info from resumes
        self.load_res.setStyleSheet(self.btn_style())# set the style to the return of the btn_style funtion (the css style sheet)
        self.load_res.clicked.connect(self.get_matches)# connect button to the get_matches function
        
        self.show_res=qtw.QPushButton("Show Results")# define a new button with a label. the button shows the tabel of results
        self.show_res.setStyleSheet(self.btn_style())# set the style to the return of the btn_style funtion (the css style sheet)
        self.show_res.clicked.connect(self.visit)# connect button to the visit function
        
        self.back=qtw.QPushButton("Back to Home page")# define a new button with a label. the button goes back to home page
        self.back.setStyleSheet(self.btn_style())# set the style to the return of the btn_style funtion (the css style sheet)
        self.back.clicked.connect(self.back_home)# connect button to the back_home function

        ###################################################
        self.skills=qtw.QLineEdit()#define a QLineEdit to add skills in
        self.skills_label=qtw.QLabel('Enter skills you want:')#define a label
        self.skills_label.setStyleSheet(self.label_style())# set the style to the return of the label_style funtion (the css style sheet)
        
        self.skills.setStyleSheet(self.line_edit_style())# set the style to the return of the line_edit_style funtion (the css style sheet)
        self.skills.editingFinished.connect(self.add_word)# connect Line edit to the add_word function(when clicking enter while in the text box)
        
        
        self.clear=qtw.QPushButton("Clear")# define a new button with a label. the button clears the list of skills
        self.clear.clicked.connect(self.clear_list)# cnnect button to the clear_list function
        self.clear.setStyleSheet(self.btn_style())# set the style to the return of the btn_style funtion (the css style sheet)
        
        self.remove_keyword=qtw.QPushButton('remove')# define a new button with a label . the button removes the selected item fromthe skill list
        self.remove_keyword.clicked.connect(self.remove_word)# connect button to the remove_word function
        self.remove_keyword.setStyleSheet(self.btn_style())# set the style to the return of the btn_style funtion (the css style sheet)
        
        self.skills_list=qtw.QListWidget()# creat the list widget where we will ad the skills
        
        self.add_keyword=qtw.QPushButton('Add skill')# define a new button with a label . the button adds new skills 
        self.add_keyword.setStyleSheet(self.btn_style())# set the style to the return of the btn_style funtion (the css style sheet)
        self.add_keyword.clicked.connect(self.add_word)
        
        self.save_list=qtw.QPushButton('save skills')# define a new button with a label . the button saves the skills into a .txt file
        self.save_list.setStyleSheet(self.btn_style())# set the style to the return of the btn_style funtion (the css style sheet)
        self.save_list.clicked.connect(self.save_skills)# connect button to the save_skills function
        
        self.tabel=qtw.QTableWidget(4,2)# create an empty table
        self.tabel.setColumnWidth(0,400)# set the 1st column width to 400 pixel
        self.tabel.setColumnWidth(1,100)# set the 2nd column width to 100 pixel
        
        
        self.setup_layout()# call the setup_layout function
    def save_skills(self):# define the save_skills function
        l=[]#empty list to get the skills
        for i in range(self.skills_list.count()):# iterate from 0 to num_items in the list
            l.append(self.skills_list.item(i).text())# append item to l
        l=' '.join(l)#transform l to a space seprated string
        if len(l)>0:# if not empty
            with open('skills.txt','w') as f:# save string in a text file called skills.txt
                f.write(l)#write the string into the file
                
    def get_matches(self):
        
        my_resume=resumes(self.resumes_dir,self.job_disc_dir)# creat an object of resumes class and pass the resumes folder directory and job description path 
        #we got from the QFileDialogs
        my_resume.extract_data()#call the extract data function
        matches=my_resume.get_matches()#get matches
        matches=dict(sorted(matches.items(), key=lambda item: item[1],reverse=True))# sort items in the dictionaries by value in descending order
        self.tabel.setRowCount(len(matches))# set number of rows in tabel to be length of resumes found
        for ind,i in enumerate(matches):# iterate through keys of the dictionary matches (names)
            self.tabel.setItem(ind,0,qtw.QTableWidgetItem(str(i)))# set first column at index row ind to be the name
            self.tabel.setItem(ind,1,qtw.QTableWidgetItem(str(matches[i])))# set second column at row index ind to be the matching similarity
        
    def visit(self):# define the visit function
        self.stckd.setCurrentIndex(1)# go to the tabel page
    def back_home(self):#define the go_home function
        self.stckd.setCurrentIndex(0)# got to home page
    def clear_list(self):# define the clear_list function
        self.skills_list.clear()
    def label_style(self):# define the label_style function
        return '''QLabel{
        color:white;
        font:bold 12px;
        min-width=5em;
        }
        '''# return style sheet that makes label color white makes the font bold 12px and width of 5em
    def remove_word(self):# define the remove_word function
        self.items=self.skills_list.selectedItems()#get list of all selected ites
        for item in self.items:#iterate through all selected items
            QLWI=self.skills_list.takeItem(self.skills_list.row(item))#remove each item
    def add_word(self):# define the add_word function
        self.skills_list.addItem(qtw.QListWidgetItem(self.skills.text(),self.skills_list))#ad curren text of lineedit to list
        self.skills.setText('')# clears text in the lineedit
    def line_edit_style(self):# define the line_edit_style function
        return'''
        color:white;
        min-height:2em;
        border-radius:5px;
        font:bold 14px;
        min-width:25em
        '''# return style sheet that makes LineEdit text-color white makes the font bold 14px and width of 5em and height 2em and adds curvature to the 
        #bordesrs =5px
    def btn_style(self):# define the btn_style function
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
                """# return style sheet that makes button text-color white,background-color:blue makes the font bold 14px and width of 14em and height 3em and 
                   #adds curvature to the bordesrs =5px
    def file_dialog(self):# define the file_dialog function
        res=qtw.QFileDialog.getOpenFileName(self,'Open File','C:\\',"Image files (*.pdf *.docx *.txt)")#open a filedialoge that can selct only .pdf,.docx,.txt
        self.job_disc_dir=res[0]#get the filepath
    def get_dir(self):# define the get_dir function
        file = str(qtw.QFileDialog.getExistingDirectory(self, "Select Directory"))#open a filedialouge that selcts only folders
        self.resumes_dir=file#get the directory

    def setup_layout(self):# define the setup_layout function
        
        self.stckd=qtw.QStackedLayout()#define the stacked layout that will contain all the pages of the application
        self.main_w=qtw.QWidget()# 1st page main widget
        self.home_layout=qtw.QHBoxLayout()# 1st page layout wich is a Horizontal Box Layout will contain two main vertical box layouts
        self.left_ly=qtw.QVBoxLayout()# main left layout (contains the main buttons)
        
        self.left_ly.addWidget(self.select_folder)#add button to left_ly
        self.left_ly.addWidget(self.select_file)#add button to left_ly
        self.left_ly.addWidget(self.load_res)#add button to left_ly
        self.left_ly.addWidget(self.show_res)#add button to left_ly
        self.home_layout.addLayout(self.left_ly)#add left_ly to home_layout
        self.home_layout.addStretch()# add a space between two main layouts
        
        self.right_layout=qtw.QVBoxLayout()#2nd main layout 
        self.right_layout.addWidget(self.skills_list)#add the list to the layout
        self.right_layout.addWidget(self.skills_label)#add the Label above the lineEdit to the Layout
        self.right_layout.addWidget(self.skills)# add the Line edit
        self.buttons_layout=qtw.QHBoxLayout()#create a horizontal Layout to add buttons aside
        self.buttons_layout.addWidget(self.add_keyword)# add button to buttons_layout
        self.buttons_layout.addWidget(self.remove_keyword)# add button to buttons_layout
        self.buttons_layout.addWidget(self.clear)# add button to buttons_layout
        self.right_layout.addLayout(self.buttons_layout)# add buttons_layout to right_layout
        self.right_layout.addWidget(self.save_list)## add button to right_layout
        self.home_layout.addLayout(self.right_layout)#add right_layout to home_layout
        self.main_w.setLayout(self.home_layout)#add home_layout to main_widget
        
        
        
        self.tabel_ly_w=qtw.QWidget()#create tabele page main widget
        self.tabel_ly=qtw.QVBoxLayout()##create tabele page main layout
        self.tabel_ly.addWidget(self.back)# add the back button 
        self.tabel_ly.addWidget(self.tabel)# add the tabel
        self.tabel_ly.addStretch()#add a space at the end
        self.tabel_ly_w.setLayout(self.tabel_ly)#set tabel_ly_w layout to be tabel_ly

        
        
        
        self.stckd.addWidget(self.main_w)#add home page widget stcked
        self.stckd.addWidget(self.tabel_ly_w)#add tabel page widget stcked
        self.stckd.setCurrentIndex(0)# default page to be home page
        self.setLayout(self.stckd)# set main layout of the whole application to be stckd        
    
        self.show()# show the application 
if __name__=='__main__':# if we are in main python code in the project
    app=qtw.QApplication([])# creat a qtw app
    app.setStyleSheet(qdarkstyle.load_stylesheet_pyqt5())# set mode to dark mode
    mw=MainWindow()# create an object of aMainWindow class
    app.exec_()#execute the app
   