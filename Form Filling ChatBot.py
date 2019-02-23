import speech_recognition as sr
import pyglet

from gtts import gTTS
import webbrowser as wb
import win32com.client as wincl
import time
#import speake3
import pyttsx
#For Directory Browsing
import glob
import ntpath

class VoiceRecognition:
    @staticmethod
    def InitMicAndReturnVoice():
        textFromVoice = ""
        r = sr.Recognizer()
        try :
            with sr.Microphone() as source:                                                                       
                print("Speak...")
                voiceOutput = r.listen(source,timeout=5,phrase_time_limit=5)
                textFromVoice = r.recognize_google(voiceOutput)
                time.sleep(5)
        except Exception as ex:
            textFromVoice = str(ex)
        finally:
            del r
        return textFromVoice
    @staticmethod
    def TextToVoice(msg,voiceOutput):
        print(msg + str(voiceOutput))
        lang ='en'
        engine = pyttsx.init()
        engine.setProperty('rate', 150)
        engine.say(voiceOutput,lang)
        engine.runAndWait()
##    @staticmethod
##    def DataConfirmation() :
####        TextToVoice("","Do you want to save the data (say Yes/No)")
####        textOutput = InitMicAndReturnVoice()
##        if textOutput == "yes":
##            DataConfirmationStatus = True
##        else :
##            DataConfirmationStatus = False
class Questions:
    QuestionsArray = []
    @staticmethod
    def SetDataToQuestionsArray(fileName):
        path = 'C:/questionsinput/'
        f= open(path + fileName + ".csv","r")
        for ques in f.read().split(','):
            Questions.QuestionsArray.append(ques)
        f.close()
    @staticmethod
    def GetQuestionFiles():
        path = 'C:/questionsinput/*.csv*'
        questionsArray = []
        questionFiles = glob.glob(path)
        for question in questionFiles:
            questionFullName = ntpath.basename(question)
            questionName,ext = questionFullName.split('.', 1)
            questionsArray.append(questionName)
        return questionsArray
class Confirmation :
    DataConfirmationStatus = False
    @staticmethod
    def GetUserSelectedFilename():
        files = Questions.GetQuestionFiles()
        filesString =  ', '.join([str(x) for x in files])
        VoiceRecognition.TextToVoice("Please select the question file : ","Please select the question file " + filesString)
        return VoiceRecognition.InitMicAndReturnVoice()

fileName = Confirmation.GetUserSelectedFilename()


Questions.SetDataToQuestionsArray(fileName)
#VoiceRecognition.DataConfirmation()

i = 0
f= open("dataOutput" + ".csv","w+")
while i < len(Questions.QuestionsArray):
    VoiceRecognition.TextToVoice("Question" + str(i+1) + " : " ,Questions.QuestionsArray[i])
    f.write("Question " + str(i+1) + " : " + Questions.QuestionsArray[i] + "\n")
    textFromVoice = VoiceRecognition.InitMicAndReturnVoice()
        #time.sleep(5)
    VoiceRecognition.TextToVoice("You said: " ,textFromVoice)
    f.write("Answer : " + textFromVoice + "\n")
    f.write("================" + "\n")
    i +=1
VoiceRecognition.TextToVoice("","Data inserted successfully.")
f.close()

