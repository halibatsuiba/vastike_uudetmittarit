# -*- coding: utf-8 -*-
import imaplib
import email
import unicodedata
import HTMLParser
import quopri
import datetime
import json
from email.header import decode_header

class hae_postit():
    def __init__(self):
        print "Hae_postit - init"
        
    def extract_body(self, payload): 
        if isinstance(payload,str):
            return payload
        else:
            return '\n'.join([extract_body(part.get_payload()) for part in payload])

    '''
    Etsi talon kirjainta otsikosta.
    Palauttaa "NA" jos ei loydy.
    '''
    def etsi_talotunniste(self, subject):
        tunniste = decode_header(subject)[0][0]
        if tunniste.lower().find("talo") > -1:
            try:
                taloTunniste = tunniste.split(" ")
                for item in taloTunniste:
                  if len(item) == 1 and item in "ABCDEFG":
                    return item
            except:
                return "NA"
        else:
            return "NA"
    
    '''
    Huhhei
    '''
    def get_current_timeframe(self):
        dayOfMonth = int(datetime.date.today().strftime("%d"))
        currentMonth = int(datetime.date.today().strftime("%m"))
        currentYear =  int(datetime.date.today().strftime("%Y"))

        if dayOfMonth >= 15:
          startMonth = currentMonth
          endMonth = currentMonth + 1
          startYear = currentYear
          endYear = currentYear
          if currentMonth == 12:
            endYear += 1

            
        elif dayOfMonth < 15:
          startMonth = currentMonth - 1
          endMonth = currentMonth
          endYear = currentYear
          startYear = currentYear

          if currentMonth == 1:
            startYear -= 1
            
        fromDate = datetime.date(startYear,startMonth,15).strftime("%d-%b-%Y")
        toDate =  datetime.date(endYear,endMonth,14).strftime("%d-%b-%Y")

        timeFrameStr = '(SINCE '+fromDate+' BEFORE '+toDate+')'
        kausiStr = str(startMonth).zfill(2)+str(startYear)
        
        if startMonth == 1:
            prevMonth = 12
            prevYear = startYear -1
        else:
            prevMonth = startMonth -1
            prevYear = startYear
            
        prevKausiStr = str(prevMonth).zfill(2)+str(prevYear)
        print timeFrameStr, kausiStr, prevKausiStr
        
        return timeFrameStr, kausiStr, prevKausiStr
    
    def is_number(self,s):
        try:
            float(s)
            return True
        except ValueError:
            pass
     
        try:
            import unicodedata
            unicodedata.numeric(s)
            return True
        except (TypeError, ValueError):
            pass
 
        return False    

    def main(self):
        talon_lukemat = {}
        lukemaraportti = {}

        with open('data.txt') as data_file:    
            vastikeData = json.load(data_file)

        #Avaa yhteys sahkopostiin
        palvelin = vastikeData["salasanat"]["palvelin"]
        tunnus = vastikeData["salasanat"]["tunnus"]
        salasana = vastikeData["salasanat"]["salasana"]

        conn = imaplib.IMAP4_SSL(palvelin, 993)
        conn.login(tunnus, salasana)
        conn.select()

        timeframe, raporttikausi, prevKausi = mailObj.get_current_timeframe()
        print timeframe, raporttikausi
        typ, data = conn.search(None, timeframe)

        with open('data.txt') as data_file:    
            vastikeLukemat = json.load(data_file)

        
        try:
            for taloTunniste in "ABCDEFG":
                vastikeLukemat["talot"][taloTunniste][raporttikausi] = {}
                
            for num in data[0].split():
                
                typ, msg_data = conn.fetch(num, '(RFC822)')
                for response_part in msg_data:
                    if isinstance(response_part, tuple):
                        msg = email.message_from_string(response_part[1])
                        subject=msg['subject']
                        
                        if msg.is_multipart():
                            #print "Multipart message"
                            html = None
                            for part in msg.get_payload():
                                if part.get_content_charset() is None:
                                    charset = chardet.detect(str(part))['encoding']
                                else:
                                    charset = part.get_content_charset()
                                    
                                if part.get_content_type() == 'text/plain':
                                    
                                    text = unicode(part.get_payload(decode=True),str(charset),"ignore").encode('utf8','replace')
                                    
                        else:
                            #print "Not a multipart message"
                            charset = msg.get_content_charset()
                            text = unicode(msg.get_payload(decode=True),str(charset),"ignore").encode('utf8','replace')


                        
                        #talon_lukemat["Talo"] = mailObj.handle_subject(subject)
                        taloTunniste = mailObj.etsi_talotunniste(subject)
                        

                        #kaikkiLukemat = text.rstrip().split('\r\n')
                        kaikkiLukemat = text.rstrip().split('\n')
                        talon_lukemat = {}
                        print taloTunniste, raporttikausi
                        for mittariLukema in kaikkiLukemat:
                            
                            mittariLukema = mittariLukema.rstrip('\r')
                            #print mittariLukema
                            #Jos talotunnistetta ei ollut otsikossa, etsi sita viestista
                            if taloTunniste == "NA":
                                if mittariLukema.lower().find("talo") > -1:
                                    if mittariLukema.find(" ") > -1:
                                        taloTunniste = mittariLukema.split(" ")[1]
                                        
                                    if mittariLukema.find(":") > -1 :
                                        taloTunniste = mittariLukema.split(":")[1]

                            lukema = None
                            mittari = None
                            mittariLukema = mittariLukema.strip().replace(" ","")
                            
                            if mittariLukema.find(":") > -1:
                                lukema = mittariLukema.split(":")[1].replace(",",".")
                                mittari = mittariLukema.split(":")[0]
                            elif mittariLukema.find("=") > -1:
                                print mittariLukema
                                lukema = mittariLukema.split("=")[1].replace(",",".")
                                mittari = mittariLukema.split("=")[0]

                            
                            if (mittari<>None) and (lukema<>None):
                                if not self.is_number(lukema):
                                    #print "Einumero!"
                                    lukema = 0
                                else:
                                    lukema = lukema.rstrip()

                                # Talojen mittarit
                                if mittari.lower().find('kylm') > -1:
                                    talon_lukemat["KylmaVesi"] = float(lukema)
                                  
                                if mittari.lower().find('kuum') > -1:
                                    talon_lukemat["LamminVesi"] = float(lukema)
                                  
                                if mittari.lower() ==('kierto'):
                                    talon_lukemat["KiertoVesi"] = float(lukema)
                                    
                                if mittari.lower() ==('kiertomwh'):
                                    talon_lukemat["kiertomwh"] = float(lukema)                                    
                                  
                                if mittari.lower() == ('lämpö') or mittari.lower() == ('lampo') or mittari.lower() == ('lämmitys') or mittari.lower() == ('lammitys'):
                                    lukema = lukema.replace('.','')
                                    talon_lukemat["Lammitys"] = float(lukema)
    
                                #Yhtion mittarit
                                try:
                                    testi = vastikeLukemat["yhtionmittarit"][raporttikausi]
                                except:
                                    vastikeLukemat["yhtionmittarit"][raporttikausi] = {}
                                    
                                yhtionVesi = 0
                                if mittari.lower().find('yhtionvesi') > -1:
                                    vastikeLukemat["yhtionmittarit"][raporttikausi]["vesimittari"] = float(lukema)

                                yhtionKaukolampo = 0                                  
                                if mittari.lower()=='kaukolampo' or mittari.lower()=='kaukolämpö':
                                    vastikeLukemat["yhtionmittarit"][raporttikausi]["kaukolampo"] = float(lukema)

                                #Muut taloyhtion kulut
                                try:
                                    testi = vastikeLukemat["yhtionmenot"][raporttikausi]
                                except:
                                    vastikeLukemat["yhtionmenot"][raporttikausi] = {}
                                    
                                if mittari.lower()=='sähkö' or mittari.lower()=='sahko':
                                    vastikeLukemat["yhtionmenot"][raporttikausi]["sahko"] = float(lukema)
                                    
                                if mittari.lower().find('ytv') > -1:
                                    vastikeLukemat["yhtionmenot"][raporttikausi]["ytv"] = float(lukema)

                                if mittari.lower().find('pankki') > -1:
                                    vastikeLukemat["yhtionmenot"][raporttikausi]["pankki"] = float(lukema)

                                if mittari.lower().find('hsy') > -1:
                                    vastikeLukemat["yhtionmenot"][raporttikausi]["hsy"] = float(lukema)
                                    
                                if mittari.lower().find('tilintarkastus') > -1:
                                    vastikeLukemat["yhtionmenot"][raporttikausi]["tilijavero"] = float(lukema)

                                if mittari.lower().find('kirjanpito') > -1:
                                    vastikeLukemat["yhtionmenot"][raporttikausi]["kirjanpito"] = float(lukema)
                                    
                                vastikeLukemat["yhtionmenot"][raporttikausi]["lisamaksu"] = 175.00

                if taloTunniste == "NA":
                    print "JOTAIN VIKAA VIKAA VIKAA..."
                else:
                    #Put values to table

                    vastikeLukemat["talot"][taloTunniste][raporttikausi] = talon_lukemat


                    
                #typ, response = conn.store(num, '+FLAGS', r'(\Seen)')
        finally:
            try:
                conn.close()
            except:
                pass
            conn.logout()

        print vastikeLukemat["yhtionmenot"][raporttikausi]
        print vastikeLukemat["yhtionmittarit"][raporttikausi]

        for taloTunniste in "ABCDEFG":        
            print taloTunniste, vastikeLukemat["talot"][taloTunniste][raporttikausi]
        
        #Kirjoita tiedot tiedostoon
        #with open('data.txt', 'w') as outfile:
        #   json.dump(vastikeLukemat, outfile)            

        

if __name__ == "__main__":
    mailObj = hae_postit()
    mailObj.main()
