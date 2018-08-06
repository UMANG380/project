import speech_recognition as sr
import webbrowser
import wolframalpha
import wikipedia
import time
import os
import pyperclip
import win32com.client as wincl

v=wincl.Dispatch("SAPI.SpVoice")

cl=wolframalpha.Client('2TAVAW-U6PE8XA8EA')
att=cl.query('Test/Attempt')
r=sr.Recognizer()
r.pause_threshold = 0.7
r.energy_threshold = 400

shell = wincl.Dispatch("WScript.Shell")
v.Speak('Hello! for a list of commands,please say any "keyword list"...')
print('Hello! for a list of commands,please say any "keyword list"...')

#List of commands

google = 'search for'
acad = 'academic search'
sc = 'deep search'
wkp = 'wiki page for'
rdds = 'read this text'
bkmk = 'bookmark this page'
vid = 'video for'
wtis = 'what is'
wtar = 'what are'
whis = 'who is'
whws = 'who was'
when = 'when'
where = 'where'
how = 'how'
paint = 'open paint'
lsp = 'silence please'
lsc = 'resume listening'
stoplst = 'stop listening'

while True:

    with sr.Microphone() as source:

        try:
            print("Please Speak")
            audio = r.listen(source, timeout = None)
            message = str(r.recognize_google(audio))
            print('You said :'+ message)

            if google in message:
                words = message.split()
                del words[0:2]
                st = ' '.join(words)
                print('Google Results for:' + str(st))
                url = 'http://google.com/search?q='+st
                webbrowser.open(url)
                v.speak('Google Results for:' + str(st))

            elif acad in message:
                words = message.split()
                del words[0:2]
                st = ' '.join(words)
                print('Acadeic results for: '+str(str))
                url = ' http://scholar.google.com/scholar?q='+st
                webbrowser.open(url)
                v.speak('Academic resukts for:'+str(st))

            elif wkp in message:

                 try:
                     words=message.split()
                     del words[0:3]
                     st = ' '.join(words)
                     wkpres=wikipedia.summary(st,sentences=2)

                     try:

                         print('\n' + str(wkpres) + '\n')
                         v.Speak(wkpres)

                     except UnicodeEncodeError:
                         v.Speak(wkpres)

                 except wikipedia.exceptions.DisambiguationError as e:
                     print(e.options)
                     v.Speak("Too many results for this keyword. Please be more specific and try again")
                     continue

                 except wikipedia.exceptions.PageError as e:
                     print("The page does not exist")
                     continue

            elif sc in message:
                try:
                    words=message.split()
                    del words[0:2]
                    st=' '.join(words)
                    scq=cl.query(st)
                    sca=next(scq.results).text
                    print('The answer is'+str(sca))
                    url='http://www.wolframalpha.com/input/?f='+st
                    webbrowser.open(url)
                    v.Speak('The answer is: '+str(sca))

                except StopIteration:
                    print('Your question is ambiguous.Please ask again')
                    v.Speak('your question is ambiguous.Please ask again')

                except Exception as e:
                    print(e)
                    v.Speak(e)

            elif sav in message:
                with open('Path to your text file','a') as f:
                    f.write(pyperclip.paste())
                    print("Savin your text to file")
                    v.Speak("Saving your text file")

            elif bkmk in message:
                shell.SendKeys("d")
                v.Speak("Page Bookmarked")

            elif keywd in message:

                print( ' ')
                print('Say " ' + google + ' " to return a google search')
                print('Say " ' + acad + ' " to return a google scholar search')
                print('Say " ' + sc + ' " to return a Wolfram Alpha Query')
                print('Say " ' + wkp + ' " to return a wikipedia page')
                print('Say " ' + book + ' " to return an amazon book search')
                print('Say " ' + rdds + ' " to read the text you have highlighted and Ctrl+c(copied)')
                print('Say " ' + sav + ' " to save the text you have highlighted and ctrl+c(copied)')
                print('Say " ' + bkmk + ' " to bookmark the page you are currently reading in your browser')
                print('Say " ' + vid + ' " to return video results for your query')
                print(' For more general questions ask them naturally and I will try my best to anwer them')
                print('Say " ' + stoplast + ' " to shut down')
                print(' ')

            elif vid in message:
                words=message.split()
                st=' '.join(words)
                print('Video results for: '+str(st))
                url='https://www.youtube.com/results?search_query='+st
                webbrowser.open(url)
                v.Speak('Video results for:'+str(st))

            elif wtis in message:
               try:
                   scq=cl.query(message)
                   sca=next(scq.results).text
                   print('The answer is: '+str(sca))
                   url='http://www.wolframalpha.com/input/?f='+st
                   webbrowser.open(url)
                   v.Speak('The answer is: '+str(sca))

               except UnicodeEncodeError:

                   v.Speak('The answer is : '+str(sca))

               except StopIteration:

                   words=message.split()
                   del words[0:2]
                   st=' '.join(words)
                   print('Google results for: '+str(st))
                   url='https://google.com/search?q='+st
                   webbrowser.open(url)
                   v.Speak('Google results for: '+str(st))

        except Exception as e:
            print(e)
            v.Speak(e)
            break
        finally:
            break
