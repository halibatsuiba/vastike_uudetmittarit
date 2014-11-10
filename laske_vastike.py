from __future__ import division
import json
import datetime
import time
import sys
from hae_postit import hae_postit
import xlsxwriter



class vastikelaskelma():

    def __init__(self):
        with open('data.txt') as data_file:    
            self.vastike = json.load(data_file)
            mailObj = hae_postit()
            self.timeframe, self.kausiStr, self.prevKausiStr = mailObj.get_current_timeframe()
            self.kaikkiDatat = {'A':{},
                                'B':{},
                                'C':{},
                                'D':{},
                                'E':{},
                                'F':{},
                                'G':{},
                                'yhteensa':{}}
            
            print "Aikavali %s, kausi %s" % (self.timeframe, self.kausiStr)

    '''
    Tama funktio laskee taloyhtion asukasmaaran
    '''
    def talon_asukasluku(self, vuosi, kuukausi, talo):
        try:
            #kausi = '{0:0{width}}'.format(kuukausi, width=2)+str(vuosi)
            self.kaikkiDatat[talo]['asukasluku'] = self.vastike["talotiedot"][talo]["asukasluku"]
            return self.vastike["talotiedot"][talo]["asukasluku"]
        except:
            return -1

  
    def lammin_vesi_jyvitys(self, vuosi, kuukausi):
        print "Jyvitetaan lammin vesi..."
        asukkaita_yhteensa = 0
        for talo in "ABCDEFG":
            asukkaita_talossa = int(self.talon_asukasluku(vuosi, kuukausi, talo))
            asukkaita_yhteensa += asukkaita_talossa

        print "Asukkaita yhteensa:", asukkaita_yhteensa

        jyvitys = {}
        for talo in "ABCDEFG":
            asukkaita_talossa = int(self.talon_asukasluku(vuosi, kuukausi, talo))
            talon_lamminvesi_jyvitys = asukkaita_talossa / round(asukkaita_yhteensa,2)
            jyvitys[talo] = round(talon_lamminvesi_jyvitys,3)
            
        return jyvitys

    def yhtion_lampiman_veden_kulutus(self, vuosi, kuukausi):
        taloyhtion_mittarilukema_nyt = self.vastike["yhtionmittarit"][self.kausiStr]["vesimittari"]
        taloyhtion_mittarilukema_viimekuussa = self.vastike["yhtionmittarit"][self.prevKausiStr]["vesimittari"]

        taloyhtion_vedenkulutus = taloyhtion_mittarilukema_nyt - taloyhtion_mittarilukema_viimekuussa
        print "Yhtion vedenkulutus: ", taloyhtion_vedenkulutus
    
        kylman_veden_kulutus = 0
        for talo in "ABCDEFG":
            kylma_vesi_nyt = self.vastike["talot"][talo][self.kausiStr]["KylmaVesi"]
            kylma_vesi_viimekuussa = self.vastike["talot"][talo][self.prevKausiStr]["KylmaVesi"]
            kylman_veden_kulutus += (kylma_vesi_nyt - kylma_vesi_viimekuussa)
      
        lampiman_veden_kulutus = taloyhtion_vedenkulutus - kylman_veden_kulutus
        self.kaikkiDatat['yhteensa']['lammin_vesi'] = round((lampiman_veden_kulutus * 53) / 1000,4)
        self.kaikkiDatat['yhteensa']['kylma_vesi'] = kylman_veden_kulutus
        #print "Yhtion kylman veden kulutus: ", kylman_veden_kulutus
        #print "Yhtion lampiman veden kulutus: ", lampiman_veden_kulutus
        return lampiman_veden_kulutus

    def lampiman_veden_kulutus_per_talo(self, vuosi, kuukausi):
        print "Lasketaan lampiman veden kulutus..."
        lamminta_vetta_kulunut = self.yhtion_lampiman_veden_kulutus(vuosi, kuukausi)
        lampiman_veden_jyvitys = self.lammin_vesi_jyvitys(vuosi, kuukausi)
        lampiman_veden_laskennallinen_kulutus_per_talo = {}

        lampiman_veden_tarkistussumma = 0
        for talo in "ABCDEFG":
            lampiman_veden_laskennallinen_kulutus_per_talo[talo] = lamminta_vetta_kulunut * lampiman_veden_jyvitys[talo]
            lampiman_veden_tarkistussumma += lampiman_veden_laskennallinen_kulutus_per_talo[talo]
  
        tarkistussumma = lampiman_veden_tarkistussumma - lamminta_vetta_kulunut
        if tarkistussumma:
            print "   VIRHE!"
            print "   Tarkistussumma: %s" % (tarkistussumma)

        
        
        return lampiman_veden_laskennallinen_kulutus_per_talo

    def kylman_veden_kulutus_per_talo(self, vuosi, kuukausi):
        print "Lasketaan kylman veden kulutus..."
        kylman_veden_kulutus = {}
        
        kylma_vesi_lukema_nyt = {}
        kylma_vesi_lukema_edellinen = {}
        kylma_vesi_yhteensa = 0
        
        for talo in "ABCDEFG":
            kylma_vesi_nyt = self.vastike["talot"][talo][self.kausiStr]["KylmaVesi"]
            kylma_vesi_viimekuussa = self.vastike["talot"][talo][self.prevKausiStr]["KylmaVesi"]
            kylman_veden_kulutus[talo] = round(kylma_vesi_nyt - kylma_vesi_viimekuussa, 4)
            kylma_vesi_yhteensa += kylman_veden_kulutus[talo]
            
            kylma_vesi_lukema_nyt[talo] = kylma_vesi_nyt
            kylma_vesi_lukema_edellinen[talo] = kylma_vesi_viimekuussa
            self.kaikkiDatat[talo]['kylmavesi'] = {'lukema':kylma_vesi_nyt,
                                                   'edellinen':kylma_vesi_viimekuussa}
            
        self.kaikkiDatat['Yhtio_vesimittari'] = {'lukema':self.vastike['yhtionmittarit'][self.kausiStr]['vesimittari'],
                                           'edellinen':self.vastike['yhtionmittarit'][self.prevKausiStr]['vesimittari']}
            
            #print "%s : %s %s %s" % (talo, kylman_veden_kulutus[talo],kylma_vesi_nyt,kylma_vesi_viimekuussa)
        print kylman_veden_kulutus
            
        return kylman_veden_kulutus,kylma_vesi_lukema_nyt,kylma_vesi_lukema_edellinen, kylma_vesi_yhteensa

    def lammityksen_kulutus_per_talo(self, vuosi, kuukausi):
        print "Lasketaan lammityksen kulutus..."
        lammityksen_kulutus = {}
        lammitys_lukema = {}
        lammitys_edellinen = {}
        
        totaali_kulutus = 0
        for talo in "ABCDEFG":
            lammitys_nyt = self.vastike["talot"][talo][self.kausiStr]["Lammitys"]
            lammitys_lukema[talo] = lammitys_nyt
            
            lammitys_viimekuussa = self.vastike["talot"][talo][self.prevKausiStr]["Lammitys"]
            lammitys_edellinen[talo] = lammitys_viimekuussa
            
            lammityksen_kulutus[talo] = (lammitys_nyt - lammitys_viimekuussa) / 1000
            totaali_kulutus += lammityksen_kulutus[talo]

            self.kaikkiDatat[talo]['lammitys'] = {'lukema':lammitys_nyt,
                                                   'edellinen':lammitys_viimekuussa}
                        
                    
        self.kaikkiDatat['yhteensa']['lammitys'] = totaali_kulutus
        
        return lammityksen_kulutus, totaali_kulutus, lammitys_lukema, lammitys_edellinen

    def kierto_per_talo(self, vuosi, kuukausi):
        print "Lasketaan kierron kulutus..."
        kiertovesi = {}
        kiertovesiTotal = 0
        for talo in "ABCDEFG":
            kiertovesi_nyt = self.vastike["talot"][talo][self.kausiStr]["KiertoVesi"]
            kiertovesi_viimekuussa = self.vastike["talot"][talo][self.prevKausiStr]["KiertoVesi"]
            kiertovesi[talo] = kiertovesi_nyt - kiertovesi_viimekuussa
            if kiertovesi[talo] < 30:
                kiertovesi[talo] = 30
                kiertovesiTotal += kiertovesi[talo]
        return kiertovesi, kiertovesiTotal
    
    def kierto_mwh_per_talo(self, vuosi, kuukausi):
        print "Lasketaan kierron mwh:t..."
        kiertomwh = {}
        kiertomwhTotal = 0
        kiertomwh_lukema = {}
        kiertomwh_edellinen = {}
        
        for talo in "ABCDEFG":
            kiertomwh_nyt = self.vastike["talot"][talo][self.kausiStr]["kiertomwh"]
            kiertomwh_lukema[talo] = kiertomwh_nyt
            
            kiertomwh_viimekuussa = self.vastike["talot"][talo][self.prevKausiStr]["kiertomwh"]
            kiertomwh_edellinen[talo] = kiertomwh_viimekuussa
            
            kiertomwh[talo] = round(kiertomwh_nyt - kiertomwh_viimekuussa,4)
            kiertomwhTotal += kiertomwh[talo]
            
            self.kaikkiDatat[talo]['kiertoenergia'] = {'lukema':kiertomwh_nyt,
                                                'edellinen':kiertomwh_viimekuussa}
            
        self.kaikkiDatat['yhteensa']['kiertoenergia'] = kiertomwhTotal
        
        return kiertomwh, kiertomwhTotal, kiertomwh_lukema, kiertomwh_edellinen

    def laske_autopaikat(self, vuosi, kuukausi):
        print "Lasketaan autopaikat..."
        autopaikat = {}
        for talo in "ABCDEFG":
            autopaikat[talo] = self.vastike["talotiedot"][talo]["autopaikat"]
            self.kaikkiDatat[talo]['autopaikat'] = {'lukema':autopaikat[talo]}
            
        return autopaikat



    def hae_uusimmat_hinnat(self):
        print "Haetaan hinnat..."
        uusin_hinta = ""
        uusin_timestamp = 0
        hinnat = self.vastike["hinnat"]
        for hinta in hinnat:
            vuosi = hinta[4:8]
            kuukausi = hinta[2:4]
            paiva = hinta[0:2]
            d = datetime.datetime (int(vuosi),int(kuukausi),int(paiva))
            timestamp = time.mktime(d.timetuple())

            if timestamp > uusin_timestamp:
                uusin_timestamp = timestamp
                uusin_hinta = hinta

        hintataulukko = self.vastike["hinnat"][uusin_hinta]
        self.kaikkiDatat['hinnat'] = hintataulukko
        
        return hintataulukko

    def kaukolammon_kokonaiskulutus(self, vuosi, kuukausi):
        taloyhtion_mittarilukema_nyt = self.vastike["yhtionmittarit"][self.kausiStr]["kaukolampo"]
        print "Kaukolampo nyt",taloyhtion_mittarilukema_nyt
        taloyhtion_mittarilukema_viimekuussa = self.vastike["yhtionmittarit"][self.prevKausiStr]["kaukolampo"]
        print "Kaukolampo edellinen",taloyhtion_mittarilukema_viimekuussa
        kaukolampo_kokonaiskulutus = (taloyhtion_mittarilukema_nyt - taloyhtion_mittarilukema_viimekuussa)
        
        self.kaikkiDatat['Yhtio_kaukolampo'] = {'lukema':self.vastike['yhtionmittarit'][self.kausiStr]['kaukolampo'],
                            'edellinen':self.vastike['yhtionmittarit'][self.prevKausiStr]['kaukolampo']}        
        
        return kaukolampo_kokonaiskulutus,taloyhtion_mittarilukema_nyt,taloyhtion_mittarilukema_viimekuussa
    

    def hae_muut_yhtion_menot(self, vuosi, kuukausi):
        muut_kulut = self.vastike["yhtionmenot"][self.kausiStr]
        summa = 0
        for kulu in muut_kulut:
            print "%s: %s" % (kulu, muut_kulut[kulu])
            summa += muut_kulut[kulu]
        muutKulutPerTalo = round(summa/7,2)
        
        self.kaikkiDatat['muut_kulut'] = muut_kulut
        
        return muut_kulut, muutKulutPerTalo, summa

    def check_if_all_data_available(self):
        kaikkiDatatSaatu = True
        for talo in "ABCDEFG":
            if self.kausiStr not in self.vastike["talot"][talo]:
                print "Ei oo: ",talo
                kaikkiDatatSaatu = False
            else:
                print "On: ",talo

        return kaikkiDatatSaatu
    
    def laske_lasku(self):

        print "\n**** Aloitetaan laskelma ****"
        #laskelma = vastikelaskelma()
      
        if not self.check_if_all_data_available():
            exit()
      
      
        now = datetime.datetime.now()
    
        hinnat = self.hae_uusimmat_hinnat()
    
        #Lasketaan kulutukset
        kylmavesi_per_talo, kylmavesi_nyt, kylmavesi_edellinen,kylma_vesi_yhteensa = self.kylman_veden_kulutus_per_talo(now.year,now.month-1)
        #print kylmavesi_per_talo
    
        lamminvesi_jyvitys = self.lammin_vesi_jyvitys(now.year,now.month-1)
    
        autopaikat = self.laske_autopaikat(now.year,now.month-1)
    
        print "***** KIERTO *****"
        kierto_per_talo, kokonaiskierto = self.kierto_per_talo(now.year,now.month-1)
        print "Kierron kokonaiskulutus:", kokonaiskierto
        
        kierto_mwh_per_talo, kokonaiskierto_mwh, kiertomwh_lukema, kiertomwh_edellinen = self.kierto_mwh_per_talo(now.year,now.month-1)
        print "KiertoMWH total %sMWH" % (kokonaiskierto_mwh)
        print "Kierto per talo:"
        print kierto_mwh_per_talo
        
        print "***** MUUT KULUT *****"
        yhtion_menot, muutKulutPerTalo, muutKulutSumma = self.hae_muut_yhtion_menot(now.year,now.month-1)
        print "Muut kulut:%s per talo:%s" %(muutKulutSumma,muutKulutPerTalo)
        
        print "***** LAMMITYS *****"
        lammityksen_kulutus_per_talo, lammitys_kokonaiskulutus, lammitys_lukema, lammitys_edellinen = self.lammityksen_kulutus_per_talo(now.year,now.month-1)
        print "Taloyhtion lammityksen kokonaiskulutus:", lammitys_kokonaiskulutus
      
        yhtion_lampiman_veden_kulutus = self.yhtion_lampiman_veden_kulutus(now.year,now.month-1)
        print "Yhtion lampiman veden kulutus:",yhtion_lampiman_veden_kulutus
    
        kaukolammon_kokonaiskulutus, kaukolampo_lukema, kaukolampo_edellinen = self.kaukolammon_kokonaiskulutus(now.year,now.month-1)
        print "Kaukolammon kokonaiskulutus:", kaukolammon_kokonaiskulutus
    
        print "*** HINNAT ***"
        kaukolammon_perusmaksu = hinnat["kaukolampo_perusmaksu"]
        kaukolammon_perusmaksu_per_talo = round(kaukolammon_perusmaksu / 7,2)
        print "Kaukolammon perusmaksu: %s per talo: %s " % (kaukolammon_perusmaksu,kaukolammon_perusmaksu_per_talo)
      
        kaukolammon_yksikkohinta = hinnat["kaukolampo_yksikkohinta"]
        kaukolammon_hinta_mwh = kaukolammon_yksikkohinta * 1000
        print "Kaukolammon yksikkohinta: %s/kWh %s/mWh" % ( kaukolammon_yksikkohinta,kaukolammon_hinta_mwh)
      
        kylmavesi_hinta = hinnat["vesi"]
        print "Kylma vesi hinta:", kylmavesi_hinta
      
        kuution_lammitysenergia = hinnat["kuution_Lammitys"]
        print "Veden lammitysenergia %skWh per kuutio." %(kuution_lammitysenergia)
      
        lamminvesi_hinta = kylmavesi_hinta + ( kuution_lammitysenergia * kaukolammon_yksikkohinta)
        print "Lampiman veden kuutiohinta:", lamminvesi_hinta
      
        kaukolammon_kokonaishinta = round(kaukolammon_perusmaksu + kaukolammon_kokonaiskulutus * kaukolammon_yksikkohinta,2)
        print "Kaukolammon kokonaishinta:", kaukolammon_kokonaishinta
    
        lampiman_veden_lammitysenergia = round((yhtion_lampiman_veden_kulutus * 53) / 1000,4)
        print "Lampiman veden lammitysenergia: %s MWH" % lampiman_veden_lammitysenergia
    
        kierron_kerroin = hinnat["kierron_kerroin"]
    
        #Lasketaan hukka
        print "*** HUKKALASKU ***"
        kaukolammon_kokonaishinta = round(kaukolammon_kokonaiskulutus * kaukolammon_hinta_mwh,2)
        print "Kaukolammon kokonaishinta:",kaukolammon_kokonaishinta
      
        lammityksen_kokonaishinta = round(lammitys_kokonaiskulutus * kaukolammon_hinta_mwh,2)
        print "Lammityksen kokonaishinta:",lammityksen_kokonaishinta
      
        kierron_kokonaishinta = kokonaiskierto * kierron_kerroin
        print "Kierron kokonaishinta:",kierron_kokonaishinta

        kierron_kokonaishinta_mwh = round(kokonaiskierto_mwh * kaukolammon_hinta_mwh,2)
        print "Kierron kokonaishintaMWH:",kierron_kokonaishinta_mwh
      
        veden_lammityksen_kokonaishinta = round(lampiman_veden_lammitysenergia * kaukolammon_hinta_mwh,2)
        print "Veden lammityksen kokonaishinta:",veden_lammityksen_kokonaishinta
        
        hukka_hinta_per_talo = round((kaukolammon_kokonaishinta - (lammityksen_kokonaishinta + kierron_kokonaishinta + veden_lammityksen_kokonaishinta)) / 7,2)
        print "Hukan hinta per talo:",hukka_hinta_per_talo
        
        hukka_mwh_hinta_per_talo = round((kaukolammon_kokonaishinta - (lammityksen_kokonaishinta + veden_lammityksen_kokonaishinta + kierron_kokonaishinta_mwh))/7,2)
        print "HukkaMWH hinta per talo: ", hukka_mwh_hinta_per_talo
      
        lasku = {"A":{},"B":{},"C":{},"D":{},"E":{},"F":{},"G":{}}
    
        #Lasketaan hinnat
        for talo in "ABCDEFG":
            lasku[talo]["Lammitys"] = round(lammityksen_kulutus_per_talo[talo] * kaukolammon_hinta_mwh + kaukolammon_perusmaksu_per_talo,2)
        
            lasku[talo]["KylmaVesi"] = round(kylmavesi_per_talo[talo] * kylmavesi_hinta,2)
        
            lasku[talo]["LamminVesi"] = round(lamminvesi_jyvitys[talo] * lamminvesi_hinta * yhtion_lampiman_veden_kulutus,2)
        
            #lasku[talo]["KiertoVesi"] = round(kierto_per_talo[talo] * kierron_kerroin,2)
        
            #lasku[talo]["hukka"] = hukka_hinta_per_talo
        
            lasku[talo]["autopaikka"] = round(autopaikat[talo]*hinnat["autopaikka"],2)
            
            lasku[talo]["kiertomwh"] = round(kierto_mwh_per_talo[talo] * kaukolammon_hinta_mwh,2)
            
            lasku[talo]["hukkamwh"] = hukka_mwh_hinta_per_talo
        
            lasku[talo]["muut"] = round(muutKulutPerTalo,2)
    
    
        laskutus = {}
        laskutus['YhtionMenot'] = yhtion_menot
        #print "MUUT TALOYHTION MENOT"
        #for item in yhtion_menot:
            #print item, yhtion_menot[item], round(yhtion_menot[item]/7,2)
            
    
        #rint "\n\nTALOKOHTAISET LASKELMAT"
        taloyhtion_totaali = 0
        for talo in "ABCDEFG":
            #print lasku[talo]
            talon_totaali = 0
            laskuRivi = ""
            for item in lasku[talo]:
                laskuRivi += item + ":"+str(lasku[talo][item]) + " "
                talon_totaali += lasku[talo][item]
            #print talo, laskuRivi, talon_totaali,"\n"
            laskutus[talo] = lasku[talo]
            taloyhtion_totaali += talon_totaali
                
        print "Total:",taloyhtion_totaali
        
    
        
        

    

        # Create an new Excel file and add a worksheet.
        workbook = xlsxwriter.Workbook('vastike.xlsx')
        worksheet = workbook.add_worksheet()
        
        # Widen the first column to make the text clearer.
        #worksheet.set_column('A:A', 20)
        
        # Add a bold format to use to highlight cells.
        bold = workbook.add_format({'bold': True})

        worksheet.write('A2', 'Yhtiovastikelaskelma As. Oy Mikkelanahde ', bold)
        
        ts = time.time()
        nowString = datetime.datetime.fromtimestamp(ts).strftime('%d.%m.%Y %H:%M:%S')
        nowFormat = workbook.add_format({'num_format': 'dd.mm.yyyy hh:mm:ss'})
        worksheet.write('A3', nowString,nowFormat)

        sarakkeet = ['Talo','Kierto','Lammin vesi','Kylma vesi', 'Lammitys','Autopaikka','Hukka','Muut','Summa']
        laskuri = 0
        for sarake in sarakkeet:
            worksheet.write(5,laskuri,sarake)
            laskuri += 1

        lasku_total = 0
        rivi = 5
        for talo in "ABCDEFG":
            rivi += 1
            summa = 0
            for item in laskutus[talo]:
                summa += laskutus[talo][item]
            lasku_total += summa
            worksheet.write(rivi,0,talo)        
            worksheet.write(rivi,1,laskutus[talo]['kiertomwh'])
            worksheet.write(rivi,2,laskutus[talo]['LamminVesi'])
            worksheet.write(rivi,3,laskutus[talo]['KylmaVesi'])
            worksheet.write(rivi,4,laskutus[talo]['Lammitys'])
            worksheet.write(rivi,5,laskutus[talo]['autopaikka'])
            worksheet.write(rivi,6,laskutus[talo]['hukkamwh'])
            worksheet.write(rivi,7,laskutus[talo]['muut'])
            worksheet.write(rivi,8,summa)
            
        worksheet.write(rivi+1,7,'Total',bold)
        worksheet.write(rivi+1,8,lasku_total)
           
        # Lukemat tiedostoon
        worksheet.write('A19','Talo')
        worksheet.write('C19','Kylma vesi')
        worksheet.write('D19','Kuuma vesi')
        worksheet.write('E19','KiertoMWH')
        worksheet.write('F19','Lammitys')
        worksheet.write('G19','Henkiloluku')
        worksheet.write('H19','Autopaikat')
        
        rivi = 15
        for talo in "ABCDEFG":
            rivi += 4
            worksheet.write(rivi,0,talo)
            worksheet.write(rivi,1,'Lukema')
            worksheet.write(rivi,2,kylmavesi_nyt[talo])
            worksheet.write(rivi,4,kiertomwh_lukema[talo])
            worksheet.write(rivi,5,lammitys_lukema[talo])
            
            
            
            worksheet.write(rivi+1,1,'Edellinen lukema')
            worksheet.write(rivi+1,2,kylmavesi_edellinen[talo])
            worksheet.write(rivi+1,4,kiertomwh_edellinen[talo])
            worksheet.write(rivi+1,5,lammitys_edellinen[talo])
            
            worksheet.write(rivi+2,1,'Kulutus')
            worksheet.write(rivi+2,2,kylmavesi_per_talo[talo])
            worksheet.write(rivi+2,3,lamminvesi_jyvitys[talo])
            worksheet.write(rivi+2,4,kierto_mwh_per_talo[talo])
            worksheet.write(rivi+2,5,lammityksen_kulutus_per_talo[talo])
            worksheet.write(rivi+2,6,self.kaikkiDatat[talo]['asukasluku'])
            worksheet.write(rivi+2,7,autopaikat[talo])
            
        worksheet.write('A48','Yhteensa')
        worksheet.write('C48',kylma_vesi_yhteensa)
        worksheet.write('E48',kokonaiskierto_mwh)
        worksheet.write('F48',lammitys_kokonaiskulutus)
        

        worksheet.write('A50','Taloyhtio')
        worksheet.write('B51','Kaukolampo')
        worksheet.write('C52','Lukema')
        worksheet.write('E52',self.kaikkiDatat['Yhtio_kaukolampo']['lukema'])
        
        worksheet.write('C53','Edellinen')
        worksheet.write('E53',self.kaikkiDatat['Yhtio_kaukolampo']['edellinen'])
        
        worksheet.write('C54','Kulutus')
        worksheet.write('E54',self.kaikkiDatat['Yhtio_kaukolampo']['lukema']-self.kaikkiDatat['Yhtio_kaukolampo']['edellinen'])
        
        worksheet.write('B55','Vesimittari')
        worksheet.write('C56','Lukema')
        worksheet.write('E56',self.kaikkiDatat['Yhtio_vesimittari']['lukema'])
        worksheet.write('C57','Edellinen')
        worksheet.write('E57',self.kaikkiDatat['Yhtio_vesimittari']['edellinen'])
        worksheet.write('C58','Kulutus')
        worksheet.write('E58',self.kaikkiDatat['Yhtio_vesimittari']['lukema']-self.kaikkiDatat['Yhtio_vesimittari']['edellinen'])

        try:
            worksheet.write('B60','Muut')
        except:
            pass
        
        try:
            worksheet.write('C61','Sahko')
            worksheet.write('E61',self.kaikkiDatat['muut_kulut']['sahko'])
        except:
            pass            
            
        try:
            worksheet.write('C62','YTV')
            worksheet.write('E62',self.kaikkiDatat['muut_kulut']['ytv'])
        except:
            pass            
            
        try:    
            worksheet.write('C63','Pankki')
            worksheet.write('E63',self.kaikkiDatat['muut_kulut']['pankki'])
        except:
            pass            
            
        try:
            worksheet.write('C64','HSY jatemaksu')
            worksheet.write('E64',self.kaikkiDatat['muut_kulut']['hsy'])
        except:
            pass            
            
        try:
            worksheet.write('C65','Lisamaksu')
            worksheet.write('E65',self.kaikkiDatat['muut_kulut']['lisamaksu'])
        except:
            pass            
            
        try:
            worksheet.write('C66','Tilitarkastus + vero')
            worksheet.write('E66',self.kaikkiDatat['muut_kulut']['tilijavero'])
        except:
            pass            
            
        try:
            worksheet.write('C67','Kirjanpito + alv')
            worksheet.write('E67',self.kaikkiDatat['muut_kulut']['kirjanpito'])
        except:
            pass            
            
        try:
            worksheet.write('C68','Tontin vuokra')
            worksheet.write('E68',self.kaikkiDatat['muut_kulut']['tontinvuokra'])
        except:
            pass            

        worksheet.write('A70','Hukkalaskelma')
        worksheet.write('B71','Kaava: Hukka = Taloyhtion kaukolammon kulutus - (talojen lammitys + veden lammitys + kiertoenergia)')
        worksheet.write('B72','Taloyhtion kaukolammon kulutus')
        worksheet.write('F72',self.kaikkiDatat['Yhtio_kaukolampo']['lukema']-self.kaikkiDatat['Yhtio_kaukolampo']['edellinen'])
        worksheet.write('B73','Talojen lammitys yhteensa')
        worksheet.write('F73',self.kaikkiDatat['yhteensa']['lammitys'])
        worksheet.write('B74','Veden lammitys yhteensa')
        worksheet.write('F74',self.kaikkiDatat['yhteensa']['lammin_vesi'])
        worksheet.write('B75','Kiertoenergia yhteensa')
        worksheet.write('F75',self.kaikkiDatat['yhteensa']['kiertoenergia'])
        worksheet.write('B76','Hukkaenergia')
        worksheet.write('F76','=F72-(F73+F74+F75)')       
           
        workbook.close()

        return laskutus
      
if __name__ == "__main__":
    myObj = vastikelaskelma()
    lasku = myObj.laske_lasku()
    lasku_total = 0
    print "******************* LASKUTUS ********************"
    for talo in "ABCDEFG":
        summa = 0
        for item in lasku[talo]:
            summa += lasku[talo][item]
        lasku_total += summa
        #print talo, lasku[talo], summa
    #print lasku['YhtionMenot']
    #print lasku_total
    print myObj.kaikkiDatat
    #myObj.tee_exceli(lasku)
    
