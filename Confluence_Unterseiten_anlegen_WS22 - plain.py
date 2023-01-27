from tokenize import Name
import requests
import csv
import json
import datetime
import re
import sys
import pandas
#Authentifizierung zum Handshake
try:
  f = open("authentication.txt", "r")
  auth_token=f.read()

except:
  print("Benötigt wird ein Personal Access Token. Ihr könnt euch diesen hier generieren: https://wikis.fu-berlin.de/plugins/personalaccesstokens/usertokens.action")
  auth_token=input("Auth Token: ")
  f= open("authentication.txt","w+")
  f.write(auth_token)

#####NEU#####
#Hier wird aus der CSV ein dictionary generiert, welches jede Anfrage mit den dazugehörigen Durchgängen und weiteren Informationen hält
def dict_generieren():
    gesamtübersicht=[]
    with open ('Prüfungsverteilung_2111022_1649Uhr - Sortierung_Mailversand_abgeschlossen.csv') as csv_file:
        csv_reader=csv.DictReader(csv_file,delimiter=';')
        line_count=0
        
        vorheriger_name=""
    #Definieren der Variablen aus der CSV
        for count,row in enumerate(csv_reader):
                    abgesagt=row["Abgesagt"]
                    if "X" in abgesagt:
                        Prüfungsname=row['Prüfungsname']

                    if "X" not in abgesagt:
                        index=row["Index"]
                        Mail=row['Mail']
                        Name=row['Name']
                        Fb=row['FB']
                        ort_hk=row['Ort HK']
                        ort_nk=row["Ort NK"]
                        Prüfungsname=row['Prüfungsname']
                        LVNummer=row['LV']
                        if LVNummer=="":
                            LVNummer="/"

                        Prüfungsdauer=row['Dauer']
                        if "andere Prüfungsdauer" in Prüfungsdauer:
                            Prüfungsdauer="120 Minuten"
                        Anzahl=row['TN']
                        if Anzahl=="":
                            Anzahl="/"

                        datum_A=row['HK_Datum']
                        Uhrzeit_A=row['Uhrzeit_HK']
                        datum_B=row['NK_Datum']
                        Uhrzeit_B=row['Uhrzeit_NK']
                        wiki_unterseite_hk=row["Wiki-Unterseite HK"]
                        wiki_unterseite_nk=row["Wiki-Unterseite NK"]

                        support=row["Support"]
                        if support=="":
                            support="Nein"

                        parallel_hk=row["Parallel HK"]
                        if parallel_hk=="":
                            parallel_hk="Nein"

                        parallel_nk=row["Parallel NK"]
                        if parallel_nk=="":
                            parallel_nk="Nein"

                        gleichzeitig_andere_prüfung_hk=row["Gleichzeitig weitere Prüfung HK"]
                        if gleichzeitig_andere_prüfung_hk=="":
                            gleichzeitig_andere_prüfung_hk=None
                        gleichzeitig_andere_prüfung_nk=row["Gleichzeitig weitere Prüfung NK"]
                        if gleichzeitig_andere_prüfung_nk=="":
                            gleichzeitig_andere_prüfung_nk=None


                        if "Durchgang" in Prüfungsname:
                            gesäubert=re.sub("Durchgang .","",Prüfungsname)

                            #print(gesäubert,vorheriger_name)
                            if gesäubert==vorheriger_name:

                                if datum_A!="":
                                    durchgang_hk=[datum_A,Uhrzeit_A]
                                    schalter=False
                                    for eintrag in gesamtübersicht[-1]["Hauptprüfung"][0]["Durchgänge"]:
                                        if eintrag==durchgang_hk:
                                            schalter=True
                                    
                                    if not schalter:
                                        gesamtübersicht[-1]["Hauptprüfung"][0]["Durchgänge"].append(durchgang_hk)

                                if datum_B!="":
                                    durchgang_nk=[datum_B,Uhrzeit_B]
                                    schalter=False

                                    for eintrag in gesamtübersicht[-1]["Nachprüfung"][0]["Durchgänge"]:
                                        if eintrag==durchgang_nk:
                                            schalter=True

                                    if not schalter:
                                        gesamtübersicht[-1]["Nachprüfung"][0]["Durchgänge"].append(durchgang_nk)
                                


                            else:

                                if datum_A!="":
                                    durchgang_1=[[datum_A,Uhrzeit_A]]
                                else:
                                    durchgang_1=None
                                
                                if datum_B!="":
                                    durchgang_2=[[datum_B,Uhrzeit_B]]
                                else:
                                    durchgang_2=None

                                übersicht={"Index":index,"Prüfungsname":gesäubert,"Dozenten":Name,"TN":Anzahl,"Fachbereich":Fb,"LV-Nummer":LVNummer,"Mail":Mail,"Prüfungsdauer":Prüfungsdauer,"Ort HK":ort_hk,"Ort NK":ort_nk,"Support":support,"Parallel HK":parallel_hk,"Parallel NK":parallel_nk,"Gleichzeitig weitere Prüfung HK":gleichzeitig_andere_prüfung_hk,"Gleichzeitig weitere Prüfung NK":gleichzeitig_andere_prüfung_nk,"Wiki-Unterseite HK":wiki_unterseite_hk,"Wiki-Unterseite NK":wiki_unterseite_nk,"Lizenzname":"",
                                    "Hauptprüfung":[
                                        {"Durchgänge":
                                            durchgang_1
                                        }],
                                    "Nachprüfung":[
                                        {"Durchgänge":
                                            durchgang_2
                                        }]}
                                gesamtübersicht.append(übersicht)
                            vorheriger_name=gesäubert
                                



                        else:
                            if datum_A!="":
                                durchgang_1=[[datum_A,Uhrzeit_A]]
                            else:
                                durchgang_1=None
                            
                            if datum_B!="":
                                durchgang_2=[[datum_B,Uhrzeit_B]]
                            else:
                                durchgang_2=None
                            
                            übersicht={"Index":index,"Prüfungsname":Prüfungsname,"Dozenten":Name,"TN":Anzahl,"Fachbereich":Fb,"LV-Nummer":LVNummer,"Mail":Mail,"Prüfungsdauer":Prüfungsdauer,"Ort HK":ort_hk,"Ort NK":ort_nk,"Support":support,"Parallel HK":parallel_hk,"Parallel NK":parallel_nk,"Gleichzeitig weitere Prüfung HK":gleichzeitig_andere_prüfung_hk,"Gleichzeitig weitere Prüfung NK":gleichzeitig_andere_prüfung_nk,"Wiki-Unterseite HK":wiki_unterseite_hk,"Wiki-Unterseite NK":wiki_unterseite_nk,"Lizenzname":"",
                                "Hauptprüfung":[
                                    {"Durchgänge":
                                        durchgang_1
                                    }],
                                "Nachprüfung":[
                                    {"Durchgänge":
                                        durchgang_2
                                    }]}

                            gesamtübersicht.append(übersicht)

    return gesamtübersicht


gesamtübersicht=dict_generieren()



def unterseite_generieren(gesamtübersicht):
  
  mutterseite=1325638661

  #Variable zum Hochzählen der angelegten Seite
  x=0
  
  def seite_post(unterseitenname,unterseite_blanko):
    
      #Name des gesamten Wiki-Bereichs
      space_key = 'eexam'

      #API-Website
      url = 'https://wikis.fu-berlin.de/rest/api/content/'

      # Request Headers
      headers = {
      'Content-Type': 'application/json;charset=iso-8859-1',
      "Authorization": f"Bearer {auth_token}"
      }
      # JSON Payload wird generiert
      data = {
      'type': 'page',
      'title': unterseitenname,
      'ancestors': [{'id':mutterseite}],
      'space': {'key':space_key},
      'body': {
          'storage':{
              'value': unterseite_blanko,
              'representation':'storage',
          }
      }
      }

      #Hochzählen nach Anlegen der Seite, um beim Printen die Seitenzahl anzeigen zu können
      #x=x+1

      #API Befehl wird ausgelöst
      try:

          r = requests.post(url=url, data=json.dumps(data), headers=headers)

          # Consider any status other than 2xx an error
          if not r.status_code // 100 == 2:
              print("Error: Unexpected response {}".format(r))
              print(r.text)
          else:
              seite = 'Seite "{}" (Nr {}) angelegt'.format(unterseitenname,x)
              antwort=r.text
              antwort=json.loads(antwort)
              seiten_id=antwort["id"]
              unterseite_hk_url=f"https://wikis.fu-berlin.de/pages/viewpage.action?pageId={seiten_id}"
              return unterseite_hk_url
      except requests.exceptions.RequestException as e:

      # A serious problem happened, like an SSLError or InvalidURL
          print("Error: {}".format(e))
          print(r.text)
      
      sys.exit()
    
  for count,i in enumerate(gesamtübersicht):
    #Erstellung der HK-Seite
    name_prüfung=i["Prüfungsname"]
    dozent=i["Dozenten"]
    tn=i["TN"]
    fb=i["Fachbereich"]
    lvnummer=i["LV-Nummer"]
    if fb=="FB Veterinärmedizin":
      lvnummer="0"+lvnummer
    mail=i["Mail"]
    prüfungsdauer=i["Prüfungsdauer"]
    support=i["Support"]
    parallel=i["Parallel HK"]
    if parallel=="Ja":
      ort="EEC1 und EEC2"
    else:
      ort=i["Ort HK"]
    
    if i["Hauptprüfung"][0]["Durchgänge"]!=None:
      datum=i["Hauptprüfung"][0]["Durchgänge"][0][0]
      durchgänge_gesamt=i["Hauptprüfung"][0]["Durchgänge"]
      print(name_prüfung,durchgänge_gesamt)
      durchgänge_einzeln=""
      for eintrag in durchgänge_gesamt:
        
        durchgänge_einzeln=durchgänge_einzeln+"<li>"+eintrag[1]+" "+f"({ort})"+"</li>"

      word_vorlage=(datetime.datetime.strptime(datum, "%d.%m.%Y")-datetime.timedelta(weeks=8)).strftime("%Y-%m-%d")
      tn_vorlage=(datetime.datetime.strptime(datum, "%d.%m.%Y")-datetime.timedelta(weeks=4)).strftime("%Y-%m-%d")
      prüfungserstellung=(datetime.datetime.strptime(datum, "%d.%m.%Y")-datetime.timedelta(weeks=5)).strftime("%Y-%m-%d")
      plattformeinstellungen=(datetime.datetime.strptime(datum, "%d.%m.%Y")-datetime.timedelta(weeks=5)).strftime("%Y-%m-%d")

      fachbereiche=[{"FB Biologie, Chemie, Pharmazie":"bcp"},{"FB Erziehungswissenschaft und Psychologie":"erzpsy"},{"FB Veterinärmedizin":"vetmed"},{"FB Wirtschaftswissenschaft":"wiwiss"},{"FB Physik":"physik"},{"FB Politik- und Sozialwissenschaften":"jfk"},{"FB Geowissenschaften":"geowiss"},{"FB Politik- und Sozialwissenschaften":"polsoz"},{"FB Philosophie und Geisteswissenschaften":"philgeist"},{"ZE Sprachenzentrum":"sz"},{"FB Rechtswissenschaft":"rewiss"},{"FB Geschichts- und Kulturwissenschaften":"geschkult"},{"FB Mathematik und Informatik":"matheinf"}]

      for paare in fachbereiche:
        for key,item in paare.items():
          if fb == key:
            fb_kurz=item


      lizenzame=f"EEC-{fb_kurz}-{lvnummer}-22-wise ({name_prüfung})"
      gesamtübersicht[count]["Lizenzname"]=lizenzame

      katalogname_hk=f"{fb_kurz}-{lvnummer}-22-wise-hk ({name_prüfung})"
      katalogname_nk=f"{fb_kurz}-{lvnummer}-22-wise-nk ({name_prüfung})"
      ###Für Fullsupport###
      if support!="S":
        unterseitenname=datetime.datetime.strptime(datum, "%d.%m.%Y").strftime("%Y-%m-%d")+" "+name_prüfung+" HK "+lvnummer+f" ({dozent})"

        unterseite_blanko=f"""
          <ac:layout>
            <ac:layout-section ac:type="three_equal">
              <ac:layout-cell>
                <h1 class="auto-cursor-target">Informationen zur Prüfung</h1>
                <table class="wrapped relative-table" style="width: 100.0%;">
                  <tbody>
                    <tr>
                      <th colspan="3">allgemeines zur Prüfung</th>
                    </tr>
                    <tr>
                      <td>LV-Nr.</td>
                      <td colspan="2">
                        <div class="content-wrapper">
                          <p>{lvnummer}</p>
                        </div>
                      </td>
                    </tr>
                    <tr>
                      <td>Name der Prüfung</td>
                      <td colspan="2">
                        <div class="content-wrapper">
                          <p>{name_prüfung}</p>
                        </div>
                      </td>
                    </tr>
                    <tr>
                      <td>Prüfer:in</td>
                      <td colspan="2">
                        <div class="content-wrapper">
                          <p>{dozent}, {mail}</p>
                        </div>
                      </td>
                    </tr>
                    <tr>
                      <td>Datum + Uhrzeit</td>
                      <td colspan="2">
                        <div class="content-wrapper">
                          <p>{datum}</p>
                          <p>Durchgänge:</p>
                          <ul>
                            {durchgänge_einzeln}
                          </ul>
                        </div>
                      </td>
                    </tr>
                    <tr>
                      <td>Dauer</td>
                      <td colspan="2">
                        <div class="content-wrapper">
                          <p>{prüfungsdauer}</p>
                        </div>
                      </td>
                    </tr>
                    <tr>
                      <td>HK/NK</td>
                      <td colspan="2">
                        <p>Link zur HK/NK:</p>
                      </td>
                    </tr>
                    <tr>
                      <td>Ort der Prüfung</td>
                      <td colspan="2">
                        <div class="content-wrapper">
                          <p>{ort}</p>
                        </div>
                      </td>
                    </tr>
                    <tr>
                      <td>Safe Exam Browser</td>
                      <td colspan="2">
                        <div class="content-wrapper">
                          <p>
                            <br/>
                          </p>
                        </div>
                      </td>
                    </tr>
                  </tbody>
                  <colgroup> <col style="width: 23.7611%;"/> <col style="width: 30.0085%;"/> <col style="width: 46.2861%;"/> </colgroup>
                  <tbody>
                    <tr>
                      <th colspan="3">Informationen zur Durchführung</th>
                    </tr>
                    <tr>
                      <td>Webex-Raum</td>
                      <td colspan="2">
                        <br/>
                      </td>
                    </tr>
                    <tr>
                      <td>Uhrzeit zum Erscheinen</td>
                      <td colspan="2">
                        <p>30 Minuten vor Prüfungsbeginn</p>
                      </td>
                    </tr>
                    <tr>
                      <td>Besonderheiten</td>
                      <td colspan="2">
                        <div class="content-wrapper">
                          <p>
                            <br/>
                          </p>
                        </div>
                      </td>
                    </tr>
                  </tbody>
                </table>
                <p class="auto-cursor-target">
                  <br/>
                </p>
                <table class="relative-table wrapped" style="width: 100.0%;">
                  <tbody>
                    <tr>
                      <th colspan="3">Informationen für Studierende</th>
                    </tr>
                    <tr>
                      <td>Login</td>
                      <td colspan="2">
                        <br/>
                      </td>
                    </tr>
                    <tr>
                      <td>N Studis</td>
                      <td colspan="2">
                          <p>{tn}</p>
                      </td>
                    </tr>
                    <tr>
                      <td>Nachteilsausgleiche</td>
                      <td colspan="2">
                        <br/>
                      </td>
                    </tr>
                    <tr>
                      <td colspan="1">Nachmeldungen</td>
                      <td colspan="2">
                        <br/>
                      </td>
                    </tr>
                  </tbody>
                  <colgroup> <col style="width: 30.86%;"/> <col style="width: 5.75127%;"/> <col style="width: 63.4444%;"/> </colgroup>
                </table>
                <p class="auto-cursor-target">
                  <br/>
                </p>
              </ac:layout-cell>
              <ac:layout-cell>
                <h1 class="auto-cursor-target">Dokumente und Deadlines</h1>
                <table class="relative-table wrapped" style="width: 100.0%;">
                  <colgroup> <col style="width: 17.1451%;"/> <col style="width: 10.8261%;"/> <col style="width: 72.0846%;"/> </colgroup>
                  <tbody>
                    <tr>
                      <th colspan="3">Deadlines</th>
                    </tr>
                    <tr>
                      <td>Wordvorlage</td>
                      <td colspan="2">
                        <div class="content-wrapper">
                          <ac:task-list>
          <ac:task>
          <ac:task-id>18</ac:task-id>
          <ac:task-status>incomplete</ac:task-status>
          <ac:task-body>
                                <ac:link>
                                  <ri:user ri:userkey="20ad2a7e80924e4c018093417b7b0005"/>
                                </ac:link> Word-Vorlage bis zum <time datetime="{word_vorlage}"/> </ac:task-body>
          </ac:task>
          </ac:task-list>
                        </div>
                      </td>
                    </tr>
                    <tr>
                      <td>RM</td>
                      <td colspan="2">
                        <div class="content-wrapper">
                          <ac:task-list>
          <ac:task>
          <ac:task-id>20</ac:task-id>
          <ac:task-status>incomplete</ac:task-status>
          <ac:task-body>
                                <ac:link>
                                  <ri:user ri:userkey="20ad2a7e7dd62460017de19e99010007"/>
                                  <ac:plain-text-link-body><![CDATA[E-Examinations Rückmeldung]]></ac:plain-text-link-body>
                                </ac:link>  RM bis zum <time datetime="9999-01-01"/> </ac:task-body>
          </ac:task>
          </ac:task-list>
                        </div>
                      </td>
                    </tr>
                    <tr>
                      <td>TN</td>
                      <td colspan="2">
                        <div class="content-wrapper">
                          <ac:task-list>
          <ac:task>
          <ac:task-id>22</ac:task-id>
          <ac:task-status>incomplete</ac:task-status>
          <ac:task-body>
                                <ac:link>
                                  <ri:user ri:userkey="20ad2a7e80924e4c018093417b7b0005"/>
                                </ac:link> TN-Vorlage bis zum <time datetime="{tn_vorlage}"/> </ac:task-body>
          </ac:task>
          </ac:task-list>
                        </div>
                      </td>
                    </tr>
                  </tbody>
                </table>
                <p class="auto-cursor-target">
                  <br/>
                </p>
                <table class="relative-table wrapped">
                  <colgroup> <col style="width: 11.5423%;"/> <col style="width: 88.516%;"/> </colgroup>
                  <tbody>
                    <tr>
                      <th colspan="2">LPLUS</th>
                    </tr>
                    <tr>
                      <td>Lizenz</td>
                      <td>
                        <p>{lizenzame}</p>
                      </td>
                    </tr>
                    <tr>
                      <td>Katalog</td>
                      <td>
                        <p>{katalogname_hk}</p>
                      </td>
                    </tr>
                    <tr>
                      <td colspan="1">Randomisierung<br/>der Aufgaben</td>
                      <td colspan="1">
                        <br/>
                      </td>
                    </tr>
                    <tr>
                      <td colspan="1">Prüfungseinsicht</td>
                      <td colspan="1">
                        <div class="content-wrapper">
                          <p>
                            <br/>
                          </p>
                        </div>
                      </td>
                    </tr>
                  </tbody>
                </table>
                <p class="auto-cursor-target">
                  <br/>
                </p>
                <table class="relative-table wrapped" style="width: 100.0%;">
                  <colgroup> <col style="width: 17.1557%;"/> <col style="width: 21.2008%;"/> </colgroup>
                  <tbody>
                    <tr>
                      <th colspan="2">Anhänge</th>
                    </tr>
                    <tr>
                      <td>
                        <div class="content-wrapper">
                          <p>Wordvorlage</p>
                        </div>
                      </td>
                      <td>
                        <br/>
                      </td>
                    </tr>
                  </tbody>
                  <tbody>
                    <tr>
                      <td>TN-Liste</td>
                      <td>
                        <br/>
                      </td>
                    </tr>
                    <tr>
                      <td>Änderungswünsche</td>
                      <td>
                        <br/>
                      </td>
                    </tr>
                  </tbody>
                </table>
                <p>
                  <br/>
                </p>
                <p class="auto-cursor-target">
                  <br/>
                </p>
              </ac:layout-cell>
              <ac:layout-cell>
                <h1 class="auto-cursor-target">Dispatching</h1>
                <table class="relative-table wrapped" style="width: 100.0%;">
                  <colgroup> <col style="width: 27.827%;"/> <col style="width: 29.8963%;"/> <col style="width: 42.3119%;"/> </colgroup>
                  <tbody>
                    <tr>
                      <th>To Do</th>
                      <th>Wer?</th>
                      <th>E-Mail verschickt</th>
                    </tr>
                    <tr>
                      <td>
                        <div class="content-wrapper">
                          <p>Prüfungserstellung</p>
                        </div>
                      </td>
                      <td>
                        <div class="content-wrapper">
                          <ac:task-list>
          <ac:task>
          <ac:task-id>23</ac:task-id>
          <ac:task-status>incomplete</ac:task-status>
          <ac:task-body>
                                <span> </span>
                              </ac:task-body>
          </ac:task>
          </ac:task-list>
                        </div>
                      </td>
                      <td>
                        <div class="content-wrapper">
                          <ac:task-list>
          <ac:task>
          <ac:task-id>25</ac:task-id>
          <ac:task-status>incomplete</ac:task-status>
          <ac:task-body>
                                <span> </span>
                              </ac:task-body>
          </ac:task>
          </ac:task-list>
                        </div>
                      </td>
                    </tr>
                    <tr>
                      <td>QK</td>
                      <td>
                        <ac:task-list>
          <ac:task>
          <ac:task-id>92</ac:task-id>
          <ac:task-status>incomplete</ac:task-status>
          <ac:task-body> </ac:task-body>
          </ac:task>
          </ac:task-list>
                      </td>
                      <td>
                        <ac:task-list>
          <ac:task>
          <ac:task-id>28</ac:task-id>
          <ac:task-status>incomplete</ac:task-status>
          <ac:task-body>
                              <span> </span>
                            </ac:task-body>
          </ac:task>
          </ac:task-list>
                      </td>
                    </tr>
                    <tr>
                      <td>TN-Liste</td>
                      <td>
                        <ac:task-list>
          <ac:task>
          <ac:task-id>93</ac:task-id>
          <ac:task-status>incomplete</ac:task-status>
          <ac:task-body> </ac:task-body>
          </ac:task>
          </ac:task-list>
                      </td>
                      <td>
                        <ac:task-list>
          <ac:task>
          <ac:task-id>31</ac:task-id>
          <ac:task-status>incomplete</ac:task-status>
          <ac:task-body>
                              <span> </span>
                            </ac:task-body>
          </ac:task>
          </ac:task-list>
                      </td>
                    </tr>
                    <tr>
                      <td>Änderungswünsche</td>
                      <td>
                        <ac:task-list>
          <ac:task>
          <ac:task-id>94</ac:task-id>
          <ac:task-status>incomplete</ac:task-status>
          <ac:task-body> </ac:task-body>
          </ac:task>
          </ac:task-list>
                      </td>
                      <td>
                        <ac:task-list>
          <ac:task>
          <ac:task-id>34</ac:task-id>
          <ac:task-status>incomplete</ac:task-status>
          <ac:task-body>
                              <span> </span>
                            </ac:task-body>
          </ac:task>
          </ac:task-list>
                      </td>
                    </tr>
                    <tr>
                      <td colspan="3">
                        <div class="content-wrapper">
                          <ac:task-list>
          <ac:task>
          <ac:task-id>38</ac:task-id>
          <ac:task-status>incomplete</ac:task-status>
          <ac:task-body>Prüfung final abgenommen </ac:task-body>
          </ac:task>
          </ac:task-list>
                        </div>
                      </td>
                    </tr>
                  </tbody>
                </table>
                <p class="auto-cursor-target">
                  <br/>
                </p>
                <table class="relative-table wrapped">
                  <colgroup> <col style="width: 24.114%;"/> <col style="width: 75.9444%;"/> </colgroup>
                  <tbody>
                    <tr>
                      <th colspan="2">Besonderheiten und Anmerkungen</th>
                    </tr>
                    <tr>
                      <td>Prüfungserstellung</td>
                      <td>
                        <br/>
                      </td>
                    </tr>
                    <tr>
                      <td>QK</td>
                      <td>
                        <br/>
                      </td>
                    </tr>
                    <tr>
                      <td>TN-Liste</td>
                      <td>
                        <br/>
                      </td>
                    </tr>
                    <tr>
                      <td>Änderungswünsche</td>
                      <td>
                        <br/>
                      </td>
                    </tr>
                  </tbody>
                </table>
                <p class="auto-cursor-target">
                  <br/>
                </p>
                <table class="wrapped">
                  <colgroup> <col style="width: 404.0px;"/> <col style="width: 90.0px;"/> <col style="width: 29.0px;"/> </colgroup>
                  <tbody>
                    <tr>
                      <td class="highlight-#f4f5f7" colspan="1" data-highlight-colour="#f4f5f7" title="Hintergrundfarbe : Hellgrau 100 %">
                        <strong title="">Prüfungserstellung &amp; Qualitätskontrolle</strong>
                      </td>
                      <td class="highlight-#f4f5f7" colspan="1" data-highlight-colour="#f4f5f7" title="Hintergrundfarbe : Hellgrau 100 %">
                        <strong title="">Erstellung</strong>
                      </td>
                      <td class="highlight-#f4f5f7" colspan="1" data-highlight-colour="#f4f5f7" title="Hintergrundfarbe : Hellgrau 100 %">
                        <strong title="">OK</strong>
                      </td>
                    </tr>
                    <tr>
                      <td colspan="1">
                        <strong>Editor</strong>
                      </td>
                      <td>
                        <br/>
                      </td>
                      <td>
                        <br/>
                      </td>
                    </tr>
                    <tr>
                      <td colspan="1">
                        <p>(optional) Ordinal-Skalen</p>
                      </td>
                      <td>
                        <ac:task-list>
          <ac:task>
          <ac:task-id>95</ac:task-id>
          <ac:task-status>incomplete</ac:task-status>
          <ac:task-body> </ac:task-body>
          </ac:task>
          </ac:task-list>
                      </td>
                      <td>
                        <ac:task-list>
          <ac:task>
          <ac:task-id>96</ac:task-id>
          <ac:task-status>incomplete</ac:task-status>
          <ac:task-body> </ac:task-body>
          </ac:task>
          </ac:task-list>
                      </td>
                    </tr>
                    <tr>
                      <td colspan="1">
                        <p>(optional) wurde Layout übernommen und Punkte entfernt?</p>
                      </td>
                      <td>
                        <ac:task-list>
          <ac:task>
          <ac:task-id>97</ac:task-id>
          <ac:task-status>incomplete</ac:task-status>
          <ac:task-body> </ac:task-body>
          </ac:task>
          </ac:task-list>
                      </td>
                      <td>
                        <ac:task-list>
          <ac:task>
          <ac:task-id>98</ac:task-id>
          <ac:task-status>incomplete</ac:task-status>
          <ac:task-body> </ac:task-body>
          </ac:task>
          </ac:task-list>
                      </td>
                    </tr>
                    <tr>
                      <td colspan="1">
                        <p>Bepunktung kontrolliert</p>
                      </td>
                      <td colspan="1">
                        <ac:task-list>
          <ac:task>
          <ac:task-id>99</ac:task-id>
          <ac:task-status>incomplete</ac:task-status>
          <ac:task-body> </ac:task-body>
          </ac:task>
          </ac:task-list>
                      </td>
                      <td colspan="1">
                        <ac:task-list>
          <ac:task>
          <ac:task-id>100</ac:task-id>
          <ac:task-status>incomplete</ac:task-status>
          <ac:task-body> </ac:task-body>
          </ac:task>
          </ac:task-list>
                      </td>
                    </tr>
                    <tr>
                      <td colspan="1">
                        <p>Multiple-Choice Aufgaben überprüft</p>
                      </td>
                      <td colspan="1">
                        <ac:task-list>
          <ac:task>
          <ac:task-id>101</ac:task-id>
          <ac:task-status>incomplete</ac:task-status>
          <ac:task-body> </ac:task-body>
          </ac:task>
          </ac:task-list>
                      </td>
                      <td colspan="1">
                        <ac:task-list>
          <ac:task>
          <ac:task-id>102</ac:task-id>
          <ac:task-status>incomplete</ac:task-status>
          <ac:task-body> </ac:task-body>
          </ac:task>
          </ac:task-list>
                      </td>
                    </tr>
                    <tr>
                      <td colspan="1">
                        <p>Bepunktung kontrolliert</p>
                      </td>
                      <td colspan="1">
                        <ac:task-list>
          <ac:task>
          <ac:task-id>103</ac:task-id>
          <ac:task-status>incomplete</ac:task-status>
          <ac:task-body> </ac:task-body>
          </ac:task>
          </ac:task-list>
                      </td>
                      <td colspan="1">
                        <ac:task-list>
          <ac:task>
          <ac:task-id>104</ac:task-id>
          <ac:task-status>incomplete</ac:task-status>
          <ac:task-body> </ac:task-body>
          </ac:task>
          </ac:task-list>
                      </td>
                    </tr>
                    <tr>
                      <td colspan="1">
                        <p>Wiki-Link eingefügt</p>
                      </td>
                      <td colspan="1">
                        <ac:task-list>
          <ac:task>
          <ac:task-id>105</ac:task-id>
          <ac:task-status>incomplete</ac:task-status>
          <ac:task-body> </ac:task-body>
          </ac:task>
          </ac:task-list>
                      </td>
                      <td colspan="1">
                        <ac:task-list>
          <ac:task>
          <ac:task-id>106</ac:task-id>
          <ac:task-status>incomplete</ac:task-status>
          <ac:task-body> </ac:task-body>
          </ac:task>
          </ac:task-list>
                      </td>
                    </tr>
                    <tr>
                      <td colspan="1">
                        <p>Abschließende Kontrolle Editor/Plattform durchgeführt</p>
                      </td>
                      <td colspan="1">
                        <ac:task-list>
          <ac:task>
          <ac:task-id>107</ac:task-id>
          <ac:task-status>incomplete</ac:task-status>
          <ac:task-body> </ac:task-body>
          </ac:task>
          </ac:task-list>
                      </td>
                      <td colspan="1">
                        <ac:task-list>
          <ac:task>
          <ac:task-id>108</ac:task-id>
          <ac:task-status>incomplete</ac:task-status>
          <ac:task-body> </ac:task-body>
          </ac:task>
          </ac:task-list>
                      </td>
                    </tr>
                    <tr>
                      <td colspan="1">
                        <strong>Plattform</strong>
                      </td>
                      <td colspan="1">
                        <p>
                          <br/>
                        </p>
                      </td>
                      <td colspan="1">
                        <p>
                          <br/>
                        </p>
                      </td>
                    </tr>
                    <tr>
                      <td colspan="1">
                        <p>Richtige Zeit eingestellt</p>
                      </td>
                      <td colspan="1">
                        <ac:task-list>
          <ac:task>
          <ac:task-id>109</ac:task-id>
          <ac:task-status>incomplete</ac:task-status>
          <ac:task-body> </ac:task-body>
          </ac:task>
          </ac:task-list>
                      </td>
                      <td colspan="1">
                        <ac:task-list>
          <ac:task>
          <ac:task-id>110</ac:task-id>
          <ac:task-status>incomplete</ac:task-status>
          <ac:task-body> </ac:task-body>
          </ac:task>
          </ac:task-list>
                      </td>
                    </tr>
                    <tr>
                      <td colspan="1">Richtige Vorlagenauswahl getroffen</td>
                      <td colspan="1">
                        <ac:task-list>
          <ac:task>
          <ac:task-id>111</ac:task-id>
          <ac:task-status>incomplete</ac:task-status>
          <ac:task-body> </ac:task-body>
          </ac:task>
          </ac:task-list>
                      </td>
                      <td colspan="1">
                        <ac:task-list>
          <ac:task>
          <ac:task-id>112</ac:task-id>
          <ac:task-status>incomplete</ac:task-status>
          <ac:task-body> </ac:task-body>
          </ac:task>
          </ac:task-list>
                      </td>
                    </tr>
                    <tr>
                      <td colspan="1">Design-Optionen hinzugefügt</td>
                      <td colspan="1">
                        <ac:task-list>
          <ac:task>
          <ac:task-id>113</ac:task-id>
          <ac:task-status>incomplete</ac:task-status>
          <ac:task-body> </ac:task-body>
          </ac:task>
          </ac:task-list>
                      </td>
                      <td colspan="1">
                        <ac:task-list>
          <ac:task>
          <ac:task-id>114</ac:task-id>
          <ac:task-status>incomplete</ac:task-status>
          <ac:task-body> </ac:task-body>
          </ac:task>
          </ac:task-list>
                      </td>
                    </tr>
                    <tr>
                      <td colspan="1">Multiple-Choice-Einstellungen gesetzt</td>
                      <td colspan="1">
                        <ac:task-list>
          <ac:task>
          <ac:task-id>115</ac:task-id>
          <ac:task-status>incomplete</ac:task-status>
          <ac:task-body> </ac:task-body>
          </ac:task>
          </ac:task-list>
                      </td>
                      <td colspan="1">
                        <ac:task-list>
          <ac:task>
          <ac:task-id>116</ac:task-id>
          <ac:task-status>incomplete</ac:task-status>
          <ac:task-body> </ac:task-body>
          </ac:task>
          </ac:task-list>
                      </td>
                    </tr>
                    <tr>
                      <td colspan="1">Reportmappe erstellt</td>
                      <td colspan="1">
                        <ac:task-list>
          <ac:task>
          <ac:task-id>117</ac:task-id>
          <ac:task-status>incomplete</ac:task-status>
          <ac:task-body> </ac:task-body>
          </ac:task>
          </ac:task-list>
                      </td>
                      <td colspan="1">
                        <ac:task-list>
          <ac:task>
          <ac:task-id>118</ac:task-id>
          <ac:task-status>incomplete</ac:task-status>
          <ac:task-body> </ac:task-body>
          </ac:task>
          </ac:task-list>
                      </td>
                    </tr>
                    <tr>
                      <td colspan="1">Dozierende freigeschaltet</td>
                      <td colspan="1">
                        <ac:task-list>
          <ac:task>
          <ac:task-id>119</ac:task-id>
          <ac:task-status>incomplete</ac:task-status>
          <ac:task-body> </ac:task-body>
          </ac:task>
          </ac:task-list>
                      </td>
                      <td colspan="1">
                        <ac:task-list>
          <ac:task>
          <ac:task-id>120</ac:task-id>
          <ac:task-status>incomplete</ac:task-status>
          <ac:task-body> </ac:task-body>
          </ac:task>
          </ac:task-list>
                      </td>
                    </tr>
                    <tr>
                      <td colspan="1">Dozierende in Nachbewertung eingetragen</td>
                      <td colspan="1">
                        <ac:task-list>
          <ac:task>
          <ac:task-id>121</ac:task-id>
          <ac:task-status>incomplete</ac:task-status>
          <ac:task-body> </ac:task-body>
          </ac:task>
          </ac:task-list>
                      </td>
                      <td colspan="1">
                        <ac:task-list>
          <ac:task>
          <ac:task-id>122</ac:task-id>
          <ac:task-status>incomplete</ac:task-status>
          <ac:task-body> </ac:task-body>
          </ac:task>
          </ac:task-list>
                      </td>
                    </tr>
                    <tr>
                      <td colspan="1">Reihenfolge der Fächer kontrolliert</td>
                      <td colspan="1">
                        <ac:task-list>
          <ac:task>
          <ac:task-id>123</ac:task-id>
          <ac:task-status>incomplete</ac:task-status>
          <ac:task-body> </ac:task-body>
          </ac:task>
          </ac:task-list>
                      </td>
                      <td colspan="1">
                        <ac:task-list>
          <ac:task>
          <ac:task-id>124</ac:task-id>
          <ac:task-status>incomplete</ac:task-status>
          <ac:task-body> </ac:task-body>
          </ac:task>
          </ac:task-list>
                      </td>
                    </tr>
                  </tbody>
                </table>
                <p>
                  <br/>
                </p>
                <p>
                  <br/>
                </p>
              </ac:layout-cell>
            </ac:layout-section>
          </ac:layout>
        """

      ###Für Support-Projekt###
      else:
        unterseitenname=datetime.datetime.strptime(datum, "%d.%m.%Y").strftime("%Y-%m-%d")+" [S] "+name_prüfung+" HK "+lvnummer+f" ({dozent})"
        
        unterseite_blanko=f"""
            <ac:layout>
              <ac:layout-section ac:type="three_equal">
                <ac:layout-cell>
                  <h1 class="auto-cursor-target">Allgemeine Informationen zur Prüfung</h1>
                  <table class="wrapped relative-table" style="width: 100.0%;">
                    <tbody>
                      <tr>
                        <th colspan="3">allgemeines zur Prüfung</th>
                      </tr>
                      <tr>
                        <td>LV-Nr.</td>
                        <td colspan="2">
                          <div class="content-wrapper">
                            <p>{lvnummer}</p>
                          </div>
                        </td>
                      </tr>
                      <tr>
                        <td>Name der Prüfung</td>
                        <td colspan="2">
                          <div class="content-wrapper">
                            <p>{name_prüfung}</p>
                          </div>
                        </td>
                      </tr>
                      <tr>
                        <td>Prüfer:in</td>
                        <td colspan="2">
                          <p>{dozent}, {mail}</p>
                        </td>
                      </tr>
                      <tr>
                        <td>Datum + Uhrzeit</td>
                        <td colspan="2">
                          <div class="content-wrapper">
                            <p>{datum}</p>
                            <p>Durchgänge:</p>
                            <ul>
                              {durchgänge_einzeln}
                            </ul>
                          </div>
                        </td>
                      </tr>
                      <tr>
                        <td>Dauer</td>
                        <td colspan="2">
                          <div class="content-wrapper">
                            <p>{prüfungsdauer}</p>
                          </div>
                        </td>
                      </tr>
                      <tr>
                        <td colspan="1">Repository Seite</td>
                        <td colspan="2">
                          <br/>
                        </td>
                      </tr>
                      <tr>
                        <td>HK/NK</td>
                        <td colspan="2">
                          <p>Link zur HK/NK:</p>
                        </td>
                      </tr>
                      <tr>
                        <td>Ort der Prüfung</td>
                        <td colspan="2">
                          <div class="content-wrapper">
                            <p>{ort}</p>
                          </div>
                        </td>
                      </tr>
                      <tr>
                        <td>Safe Exam Browser</td>
                        <td colspan="2">
                          <div class="content-wrapper">
                            <p>
                              <span class="templateparameter">nein</span>
                            </p>
                          </div>
                        </td>
                      </tr>
                    </tbody>
                    <colgroup> <col style="width: 23.7903%;"/> <col style="width: 29.974%;"/> <col style="width: 46.2914%;"/> </colgroup>
                    <tbody>
                      <tr>
                        <th colspan="3">Informationen zur Durchführung</th>
                      </tr>
                      <tr>
                        <td>Webex-Raum</td>
                        <td colspan="2">
                          <div class="content-wrapper">
                            <p>
                              <br/>
                            </p>
                          </div>
                        </td>
                      </tr>
                      <tr>
                        <td>Uhrzeit zum Erscheinen</td>
                        <td colspan="2">30 Minuten vor der Prüfung</td>
                      </tr>
                      <tr>
                        <td>Besonderheiten</td>
                        <td colspan="2">
                          <div class="content-wrapper">
                            <p>
                              <br/>
                            </p>
                          </div>
                        </td>
                      </tr>
                    </tbody>
                  </table>
                  <p class="auto-cursor-target">
                    <br/>
                  </p>
                  <table class="relative-table wrapped" style="width: 100.0%;">
                    <tbody>
                      <tr>
                        <th colspan="3">Informationen für Studierende</th>
                      </tr>
                      <tr>
                        <td>Login</td>
                        <td colspan="2">
                          <br/>
                        </td>
                      </tr>
                      <tr>
                        <td>N Studis</td>
                        <td colspan="2">
                          <p>{tn}</p>
                        </td>
                      </tr>
                      <tr>
                        <td>Nachteilsausgleiche</td>
                        <td colspan="2">
                          <p>
                            <br/>
                          </p>
                        </td>
                      </tr>
                      <tr>
                        <td colspan="1">Nachmeldungen</td>
                        <td colspan="2">
                          <p>
                            <br/>
                          </p>
                        </td>
                      </tr>
                    </tbody>
                    <colgroup> <col style="width: 30.86%;"/> <col style="width: 5.75127%;"/> <col style="width: 63.4444%;"/> </colgroup>
                  </table>
                  <p class="auto-cursor-target">
                    <br/>
                  </p>
                </ac:layout-cell>
                <ac:layout-cell>
                  <h1 class="auto-cursor-target">Dokumente und Deadlines</h1>
                  <table class="relative-table wrapped" style="width: 100.0%;">
                    <colgroup> <col style="width: 32.5762%;"/> <col style="width: 3.29784%;"/> <col style="width: 64.0589%;"/> </colgroup>
                    <tbody>
                      <tr>
                        <th colspan="3">Deadlines</th>
                      </tr>
                      <tr>
                        <td>
                          <p>Erstellung (Reminder ggf. automatisiert)</p>
                        </td>
                        <td colspan="2">
                          <div class="content-wrapper">
                            <ac:task-list>
            <ac:task>
            <ac:task-id>18</ac:task-id>
            <ac:task-status>incomplete</ac:task-status>
            <ac:task-body>
                                  <ac:link>
                                    <ri:user ri:userkey="20ad2a7e82f723140182f8233a530008"/>
                                  </ac:link> Lehrende Erstellung Prüfung bis zum <time datetime="{prüfungserstellung}"/> </ac:task-body>
            </ac:task>
            </ac:task-list>
                          </div>
                        </td>
                      </tr>
                      <tr>
                        <td colspan="1">Plattformeinstellungen (Reminder ggf. automatisiert)</td>
                        <td colspan="2">
                          <div class="content-wrapper">
                            <ac:task-list>
            <ac:task>
            <ac:task-id>2199</ac:task-id>
            <ac:task-status>incomplete</ac:task-status>
            <ac:task-body>
                                  <span class="placeholder-inline-tasks"> <span class="fabric"> <span> <span> <ac:link>
                                            <ri:user ri:userkey="20ad2a7e82f723140182f824006c0009"/>
                                          </ac:link>  Plattformeinstellungen in LPLUS bis zum <time datetime="{plattformeinstellungen}"/> <br/>
                                        </span> </span> </span> </span>
                                </ac:task-body>
            </ac:task>
            </ac:task-list>
                          </div>
                        </td>
                      </tr>
                      <tr>
                        <td>Rückmeldung zu Änderungswünschen und Prüfungsabnahme von Lehrenden</td>
                        <td colspan="2">
                          <div class="content-wrapper">
                            <ac:task-list>
            <ac:task>
            <ac:task-id>2215</ac:task-id>
            <ac:task-status>incomplete</ac:task-status>
            <ac:task-body>
                                  <span class="placeholder-inline-tasks"> <ac:link>
                                      <ri:user ri:userkey="20ad2a7e82f723140182f82586ca000b"/>
                                      <ac:plain-text-link-body><![CDATA[Lehrende-Rückmeldung]]></ac:plain-text-link-body>
                                    </ac:link> RM bis zum <time datetime="9999-01-01"/> <br/>
                                  </span>
                                </ac:task-body>
            </ac:task>
            </ac:task-list>
                          </div>
                        </td>
                      </tr>
                      <tr>
                        <td>TN</td>
                        <td colspan="2">
                          <div class="content-wrapper">
                            <ac:task-list>
            <ac:task>
            <ac:task-id>22</ac:task-id>
            <ac:task-status>incomplete</ac:task-status>
            <ac:task-body>
                                  <ac:link>
                                    <ri:user ri:userkey="20ad2a7e80924e4c018093417b7b0005"/>
                                  </ac:link> <span style="color: rgb(23,43,77);">TN-Vorlage bis zum <time datetime="{tn_vorlage}"/> </span>
                                </ac:task-body>
            </ac:task>
            </ac:task-list>
                          </div>
                        </td>
                      </tr>
                    </tbody>
                  </table>
                  <p class="auto-cursor-target">
                    <br/>
                  </p>
                  <table class="relative-table wrapped">
                    <colgroup> <col style="width: 11.5423%;"/> <col style="width: 88.516%;"/> </colgroup>
                    <tbody>
                      <tr>
                        <th colspan="2">LPLUS</th>
                      </tr>
                      <tr>
                        <td>Lizenz</td>
                        <td>
                          <p>{lizenzame}</p>
                        </td>
                      </tr>
                      <tr>
                        <td>Katalog</td>
                        <td>
                          <p>{katalogname_hk}</p>
                        </td>
                      </tr>
                      <tr>
                        <td colspan="1">Repository</td>
                        <td colspan="1">
                          <p>
                            <br/>
                          </p>
                        </td>
                      </tr>
                      <tr>
                        <td colspan="1">Randomisierung<br/>der Aufgaben</td>
                        <td colspan="1">
                          <br/>
                        </td>
                      </tr>
                      <tr>
                        <td colspan="1">Prüfungseinsicht</td>
                        <td colspan="1">
                          <div class="content-wrapper">
                            <p>
                              <br/>
                            </p>
                          </div>
                        </td>
                      </tr>
                    </tbody>
                  </table>
                  <p class="auto-cursor-target">
                    <br/>
                  </p>
                  <table class="relative-table wrapped" style="width: 100.0%;">
                    <colgroup> <col style="width: 24.912%;"/> <col style="width: 28.2303%;"/> </colgroup>
                    <tbody>
                      <tr>
                        <th colspan="2">Anhänge</th>
                      </tr>
                      <tr>
                        <td>
                          <div class="content-wrapper">
                            <p>Plattformeinstellungen</p>
                          </div>
                        </td>
                        <td>
                          <div class="content-wrapper">
                            <p>
                              <br/>
                            </p>
                            <p>
                              <span class="placeholder-inline-tasks"> <br/>
                              </span>
                            </p>
                          </div>
                        </td>
                      </tr>
                      <tr>
                        <td>TN-Liste</td>
                        <td>
                          <div class="content-wrapper">
                            <p>
                              <br/>
                            </p>
                            <p>
                              <br/>
                            </p>
                          </div>
                        </td>
                      </tr>
                    </tbody>
                  </table>
                  <p>
                    <br/>
                  </p>
                  <p class="auto-cursor-target">
                    <br/>
                  </p>
                </ac:layout-cell>
                <ac:layout-cell>
                  <h1 class="auto-cursor-target">Dispatching</h1>
                  <table class="relative-table wrapped" style="width: 100.0%;">
                    <colgroup> <col style="width: 27.7526%;"/> <col style="width: 50.5254%;"/> <col style="width: 21.7195%;"/> </colgroup>
                    <tbody>
                      <tr>
                        <th>To Do</th>
                        <th>Wer?</th>
                        <th>E-Mail verschickt</th>
                      </tr>
                      <tr>
                        <td>QK Editor + Transfer 1</td>
                        <td>
                          <div class="content-wrapper">
                            <ac:task-list>
            <ac:task>
            <ac:task-id>2253</ac:task-id>
            <ac:task-status>incomplete</ac:task-status>
            <ac:task-body>
                                  <span> </span>
                                </ac:task-body>
            </ac:task>
            </ac:task-list>
                          </div>
                        </td>
                        <td>
                          <ac:task-list>
            <ac:task>
            <ac:task-id>28</ac:task-id>
            <ac:task-status>incomplete</ac:task-status>
            <ac:task-body>
                                <span> </span>
                              </ac:task-body>
            </ac:task>
            </ac:task-list>
                        </td>
                      </tr>
                      <tr>
                        <td>QK Editor + Transfer 2</td>
                        <td>
                          <ac:task-list>
            <ac:task>
            <ac:task-id>2256</ac:task-id>
            <ac:task-status>incomplete</ac:task-status>
            <ac:task-body>
                                <span> </span>
                              </ac:task-body>
            </ac:task>
            </ac:task-list>
                        </td>
                        <td>
                          <ac:task-list>
            <ac:task>
            <ac:task-id>31</ac:task-id>
            <ac:task-status>incomplete</ac:task-status>
            <ac:task-body>
                                <span> </span>
                              </ac:task-body>
            </ac:task>
            </ac:task-list>
                        </td>
                      </tr>
                      <tr>
                        <td>TN-Liste</td>
                        <td>
                          <ac:task-list>
            <ac:task>
            <ac:task-id>2257</ac:task-id>
            <ac:task-status>incomplete</ac:task-status>
            <ac:task-body>
                                <span> </span>
                              </ac:task-body>
            </ac:task>
            </ac:task-list>
                        </td>
                        <td>
                          <ac:task-list>
            <ac:task>
            <ac:task-id>34</ac:task-id>
            <ac:task-status>incomplete</ac:task-status>
            <ac:task-body>
                                <span> </span>
                              </ac:task-body>
            </ac:task>
            </ac:task-list>
                        </td>
                      </tr>
                      <tr>
                        <td colspan="3">
                          <div class="content-wrapper">
                            <ac:task-list>
            <ac:task>
            <ac:task-id>38</ac:task-id>
            <ac:task-status>incomplete</ac:task-status>
            <ac:task-body>Prüfung durch Lehrende final abgenommen</ac:task-body>
            </ac:task>
            </ac:task-list>
                          </div>
                        </td>
                      </tr>
                    </tbody>
                  </table>
                  <p class="auto-cursor-target">
                    <br/>
                  </p>
                  <table class="relative-table wrapped" style="width: 100.0%;">
                    <colgroup> <col style="width: 13.7495%;"/> <col style="width: 86.3089%;"/> </colgroup>
                    <tbody>
                      <tr>
                        <th colspan="2">Besonderheiten und Anmerkungen</th>
                      </tr>
                      <tr>
                        <td>Transfer 1</td>
                        <td>
                          <br/>
                        </td>
                      </tr>
                      <tr>
                        <td>Transfer 2</td>
                        <td>
                          <br/>
                        </td>
                      </tr>
                      <tr>
                        <td>TN-Liste</td>
                        <td>
                          <br/>
                        </td>
                      </tr>
                    </tbody>
                  </table>
                  <p class="auto-cursor-target">
                    <br/>
                  </p>
                  <table class="wrapped relative-table" style="width: 100.0%;">
                    <colgroup> <col style="width: 67.474%;"/> <col style="width: 16.2086%;"/> <col style="width: 16.3757%;"/> </colgroup>
                    <tbody>
                      <tr>
                        <th>Checkliste Editor QK</th>
                        <th colspan="1">Transfer 1</th>
                        <th colspan="1">Transfer 2</th>
                      </tr>
                      <tr>
                        <td>
                          <p>Schriftart prüfen (Aufgaben: Arial 12 fett, Antworten: Arial 12)</p>
                        </td>
                        <td colspan="1">
                          <ac:task-list>
            <ac:task>
            <ac:task-id>2234</ac:task-id>
            <ac:task-status>incomplete</ac:task-status>
            <ac:task-body>
                                <span> </span>
                              </ac:task-body>
            </ac:task>
            </ac:task-list>
                        </td>
                        <td colspan="1">
                          <ac:task-list>
            <ac:task>
            <ac:task-id>2235</ac:task-id>
            <ac:task-status>incomplete</ac:task-status>
            <ac:task-body>
                                <span> </span>
                              </ac:task-body>
            </ac:task>
            </ac:task-list>
                        </td>
                      </tr>
                      <tr>
                        <td>
                          <p>Hinweis-Button eingefügt → nicht bei Staatsexamina </p>
                        </td>
                        <td colspan="1">
                          <ac:task-list>
            <ac:task>
            <ac:task-id>2236</ac:task-id>
            <ac:task-status>incomplete</ac:task-status>
            <ac:task-body>
                                <span> </span>
                              </ac:task-body>
            </ac:task>
            </ac:task-list>
                        </td>
                        <td colspan="1">
                          <ac:task-list>
            <ac:task>
            <ac:task-id>2237</ac:task-id>
            <ac:task-status>incomplete</ac:task-status>
            <ac:task-body>
                                <span> </span>
                              </ac:task-body>
            </ac:task>
            </ac:task-list>
                        </td>
                      </tr>
                      <tr>
                        <td>
                          <p>Auflösung der Abbildungen überprüfen</p>
                        </td>
                        <td colspan="1">
                          <ac:task-list>
            <ac:task>
            <ac:task-id>2238</ac:task-id>
            <ac:task-status>incomplete</ac:task-status>
            <ac:task-body>
                                <span> </span>
                              </ac:task-body>
            </ac:task>
            </ac:task-list>
                        </td>
                        <td colspan="1">
                          <ac:task-list>
            <ac:task>
            <ac:task-id>2239</ac:task-id>
            <ac:task-status>incomplete</ac:task-status>
            <ac:task-body>
                                <span> </span>
                              </ac:task-body>
            </ac:task>
            </ac:task-list>
                        </td>
                      </tr>
                      <tr>
                        <td>
                          <p>Hinweistext auf Anlage im Aufgabetext ergänzt</p>
                        </td>
                        <td colspan="1">
                          <ac:task-list>
            <ac:task>
            <ac:task-id>2240</ac:task-id>
            <ac:task-status>incomplete</ac:task-status>
            <ac:task-body>
                                <span> </span>
                              </ac:task-body>
            </ac:task>
            </ac:task-list>
                        </td>
                        <td colspan="1">
                          <ac:task-list>
            <ac:task>
            <ac:task-id>2241</ac:task-id>
            <ac:task-status>incomplete</ac:task-status>
            <ac:task-body>
                                <span> </span>
                              </ac:task-body>
            </ac:task>
            </ac:task-list>
                        </td>
                      </tr>
                      <tr>
                        <td colspan="1">
                          <p>Layout prüfen (leere Zeilen, zu viel Platz zwischen Fragen und Antworten)</p>
                        </td>
                        <td colspan="1">
                          <ac:task-list>
            <ac:task>
            <ac:task-id>2242</ac:task-id>
            <ac:task-status>incomplete</ac:task-status>
            <ac:task-body>
                                <span> </span>
                              </ac:task-body>
            </ac:task>
            </ac:task-list>
                        </td>
                        <td colspan="1">
                          <ac:task-list>
            <ac:task>
            <ac:task-id>2243</ac:task-id>
            <ac:task-status>incomplete</ac:task-status>
            <ac:task-body>
                                <span> </span>
                              </ac:task-body>
            </ac:task>
            </ac:task-list>
                        </td>
                      </tr>
                      <tr>
                        <td colspan="1">
                          <p>Teilbewertung (keine Maluspunkte, richtige Teilbewertung ausgewählt)</p>
                        </td>
                        <td colspan="1">
                          <ac:task-list>
            <ac:task>
            <ac:task-id>2244</ac:task-id>
            <ac:task-status>incomplete</ac:task-status>
            <ac:task-body>
                                <span> </span>
                              </ac:task-body>
            </ac:task>
            </ac:task-list>
                        </td>
                        <td colspan="1">
                          <ac:task-list>
            <ac:task>
            <ac:task-id>2245</ac:task-id>
            <ac:task-status>incomplete</ac:task-status>
            <ac:task-body>
                                <span> </span>
                              </ac:task-body>
            </ac:task>
            </ac:task-list>
                        </td>
                      </tr>
                      <tr>
                        <td colspan="1">
                          <p>Aufgaben mit festem Format prüfen</p>
                        </td>
                        <td colspan="1">
                          <ac:task-list>
            <ac:task>
            <ac:task-id>2246</ac:task-id>
            <ac:task-status>incomplete</ac:task-status>
            <ac:task-body>
                                <span> </span>
                              </ac:task-body>
            </ac:task>
            </ac:task-list>
                        </td>
                        <td colspan="1">
                          <ac:task-list>
            <ac:task>
            <ac:task-id>2247</ac:task-id>
            <ac:task-status>incomplete</ac:task-status>
            <ac:task-body>
                                <span> </span>
                              </ac:task-body>
            </ac:task>
            </ac:task-list>
                        </td>
                      </tr>
                      <tr>
                        <td colspan="1">
                          <p>Dropdown und Lückentexte prüfen (Antwortoptionen und Bepunktung)</p>
                        </td>
                        <td colspan="1">
                          <ac:task-list>
            <ac:task>
            <ac:task-id>2248</ac:task-id>
            <ac:task-status>incomplete</ac:task-status>
            <ac:task-body>
                                <span> </span>
                              </ac:task-body>
            </ac:task>
            </ac:task-list>
                        </td>
                        <td colspan="1">
                          <ac:task-list>
            <ac:task>
            <ac:task-id>2249</ac:task-id>
            <ac:task-status>incomplete</ac:task-status>
            <ac:task-body>
                                <span> </span>
                              </ac:task-body>
            </ac:task>
            </ac:task-list>
                        </td>
                      </tr>
                      <tr>
                        <td colspan="1">
                          <p>Allgemeine Auffälligkeiten (z.B. sind die Aufgaben einheitlich erstellt?) - Bitte notieren/ Rücksprache halten</p>
                        </td>
                        <td colspan="1">
                          <ac:task-list>
            <ac:task>
            <ac:task-id>2250</ac:task-id>
            <ac:task-status>incomplete</ac:task-status>
            <ac:task-body>
                                <span> </span>
                              </ac:task-body>
            </ac:task>
            </ac:task-list>
                        </td>
                        <td colspan="1">
                          <ac:task-list>
            <ac:task>
            <ac:task-id>2251</ac:task-id>
            <ac:task-status>incomplete</ac:task-status>
            <ac:task-body>
                                <span> </span>
                              </ac:task-body>
            </ac:task>
            </ac:task-list>
                        </td>
                      </tr>
                    </tbody>
                  </table>
                  <p class="auto-cursor-target">
                    <br/>
                  </p>
                  <table class="relative-table wrapped" style="width: 100.0%;">
                    <colgroup> <col style="width: 83.3404%;"/> <col style="width: 16.718%;"/> </colgroup>
                    <tbody>
                      <tr>
                        <th colspan="1">Plattform Checkliste</th>
                        <th>QK</th>
                      </tr>
                      <tr>
                        <td colspan="1">
                          <p>Richtige Zeit eingestellt</p>
                        </td>
                        <td colspan="1">
                          <ac:task-list>
            <ac:task>
            <ac:task-id>2166</ac:task-id>
            <ac:task-status>incomplete</ac:task-status>
            <ac:task-body> </ac:task-body>
            </ac:task>
            </ac:task-list>
                        </td>
                      </tr>
                      <tr>
                        <td colspan="1">Richtige Vorlagenauswahl getroffen</td>
                        <td colspan="1">
                          <ac:task-list>
            <ac:task>
            <ac:task-id>2168</ac:task-id>
            <ac:task-status>incomplete</ac:task-status>
            <ac:task-body> </ac:task-body>
            </ac:task>
            </ac:task-list>
                        </td>
                      </tr>
                      <tr>
                        <td colspan="1">Design-Optionen hinzugefügt</td>
                        <td colspan="1">
                          <ac:task-list>
            <ac:task>
            <ac:task-id>2170</ac:task-id>
            <ac:task-status>incomplete</ac:task-status>
            <ac:task-body> </ac:task-body>
            </ac:task>
            </ac:task-list>
                        </td>
                      </tr>
                      <tr>
                        <td colspan="1">Multiple-Choice-Einstellungen gesetzt</td>
                        <td colspan="1">
                          <ac:task-list>
            <ac:task>
            <ac:task-id>2172</ac:task-id>
            <ac:task-status>incomplete</ac:task-status>
            <ac:task-body> </ac:task-body>
            </ac:task>
            </ac:task-list>
                        </td>
                      </tr>
                      <tr>
                        <td colspan="1">Reportmappe erstellt</td>
                        <td colspan="1">
                          <ac:task-list>
            <ac:task>
            <ac:task-id>2174</ac:task-id>
            <ac:task-status>incomplete</ac:task-status>
            <ac:task-body> </ac:task-body>
            </ac:task>
            </ac:task-list>
                        </td>
                      </tr>
                      <tr>
                        <td colspan="1">Dozierende freigeschaltet</td>
                        <td colspan="1">
                          <ac:task-list>
            <ac:task>
            <ac:task-id>2176</ac:task-id>
            <ac:task-status>incomplete</ac:task-status>
            <ac:task-body> </ac:task-body>
            </ac:task>
            </ac:task-list>
                        </td>
                      </tr>
                      <tr>
                        <td colspan="1">Dozierende + Sys-Admins in Nachbewertung eingetragen</td>
                        <td colspan="1">
                          <ac:task-list>
            <ac:task>
            <ac:task-id>2178</ac:task-id>
            <ac:task-status>incomplete</ac:task-status>
            <ac:task-body> </ac:task-body>
            </ac:task>
            </ac:task-list>
                        </td>
                      </tr>
                      <tr>
                        <td colspan="1">Reihenfolge der Fächer kontrolliert</td>
                        <td colspan="1">
                          <ac:task-list>
            <ac:task>
            <ac:task-id>2192</ac:task-id>
            <ac:task-status>incomplete</ac:task-status>
            <ac:task-body> </ac:task-body>
            </ac:task>
            </ac:task-list>
                        </td>
                      </tr>
                    </tbody>
                  </table>
                  <p class="auto-cursor-target">
                    <br/>
                  </p>
                </ac:layout-cell>
              </ac:layout-section>
            </ac:layout>

        
        """
      #Posten der HK-Seite im Wiki
      unterseite_hk_url=seite_post(unterseitenname,unterseite_blanko)
      gesamtübersicht[count]["Wiki-Unterseite HK"]=unterseite_hk_url
    else:
      unterseite_hk_url=""

    ###Erstellung NK-Seite####
    if i["Nachprüfung"][0]["Durchgänge"]!=None:
      parallel=i["Parallel NK"]
      if parallel=="Ja":
        ort="EEC1 und EEC2"
      else:
        ort=i["Ort HK"]
      datum=i["Nachprüfung"][0]["Durchgänge"][0][0]
      durchgänge_gesamt=i["Nachprüfung"][0]["Durchgänge"]
      durchgänge_einzeln=""
      for eintrag in durchgänge_gesamt:
        
        durchgänge_einzeln=durchgänge_einzeln+"<li>"+eintrag[1]+" "+f"({ort})"+"</li>"

      word_vorlage=(datetime.datetime.strptime(datum, "%d.%m.%Y")-datetime.timedelta(weeks=8)).strftime("%Y-%m-%d")
      tn_vorlage=(datetime.datetime.strptime(datum, "%d.%m.%Y")-datetime.timedelta(weeks=4)).strftime("%Y-%m-%d")
      prüfungserstellung=(datetime.datetime.strptime(datum, "%d.%m.%Y")-datetime.timedelta(weeks=5)).strftime("%Y-%m-%d")
      plattformeinstellungen=(datetime.datetime.strptime(datum, "%d.%m.%Y")-datetime.timedelta(weeks=5)).strftime("%Y-%m-%d")
      ###Für Fullsupport###
      if support!="S":
        unterseitenname=datetime.datetime.strptime(datum, "%d.%m.%Y").strftime("%Y-%m-%d")+" "+name_prüfung+" NK "+lvnummer+f" ({dozent})"

        unterseite_blanko=f"""
          <ac:layout>
            <ac:layout-section ac:type="three_equal">
              <ac:layout-cell>
                <h1 class="auto-cursor-target">Informationen zur Prüfung</h1>
                <table class="wrapped relative-table" style="width: 100.0%;">
                  <tbody>
                    <tr>
                      <th colspan="3">allgemeines zur Prüfung</th>
                    </tr>
                    <tr>
                      <td>LV-Nr.</td>
                      <td colspan="2">
                        <div class="content-wrapper">
                          <p>{lvnummer}</p>
                        </div>
                      </td>
                    </tr>
                    <tr>
                      <td>Name der Prüfung</td>
                      <td colspan="2">
                        <div class="content-wrapper">
                          <p>{name_prüfung}</p>
                        </div>
                      </td>
                    </tr>
                    <tr>
                      <td>Prüfer:in</td>
                      <td colspan="2">
                        <div class="content-wrapper">
                          <p>{dozent}, {mail}</p>
                        </div>
                      </td>
                    </tr>
                    <tr>
                      <td>Datum + Uhrzeit</td>
                      <td colspan="2">
                        <div class="content-wrapper">
                          <p>{datum}</p>
                          <p>Durchgänge:</p>
                          <ul>
                            {durchgänge_einzeln}
                          </ul>
                        </div>
                      </td>
                    </tr>
                    <tr>
                      <td>Dauer</td>
                      <td colspan="2">
                        <div class="content-wrapper">
                          <p>{prüfungsdauer}</p>
                        </div>
                      </td>
                    </tr>
                    <tr>
                      <td>HK/NK</td>
                      <td colspan="2">
                        <p>Link zur HK/NK:</p>
                      </td>
                    </tr>
                    <tr>
                      <td>Ort der Prüfung</td>
                      <td colspan="2">
                        <div class="content-wrapper">
                          <p>{ort}</p>
                        </div>
                      </td>
                    </tr>
                    <tr>
                      <td>Safe Exam Browser</td>
                      <td colspan="2">
                        <div class="content-wrapper">
                          <p>
                            <br/>
                          </p>
                        </div>
                      </td>
                    </tr>
                  </tbody>
                  <colgroup> <col style="width: 23.7611%;"/> <col style="width: 30.0085%;"/> <col style="width: 46.2861%;"/> </colgroup>
                  <tbody>
                    <tr>
                      <th colspan="3">Informationen zur Durchführung</th>
                    </tr>
                    <tr>
                      <td>Webex-Raum</td>
                      <td colspan="2">
                        <br/>
                      </td>
                    </tr>
                    <tr>
                      <td>Uhrzeit zum Erscheinen</td>
                      <td colspan="2">
                        <p>30 Minuten vor Prüfungsbeginn</p>
                      </td>
                    </tr>
                    <tr>
                      <td>Besonderheiten</td>
                      <td colspan="2">
                        <div class="content-wrapper">
                          <p>
                            <br/>
                          </p>
                        </div>
                      </td>
                    </tr>
                  </tbody>
                </table>
                <p class="auto-cursor-target">
                  <br/>
                </p>
                <table class="relative-table wrapped" style="width: 100.0%;">
                  <tbody>
                    <tr>
                      <th colspan="3">Informationen für Studierende</th>
                    </tr>
                    <tr>
                      <td>Login</td>
                      <td colspan="2">
                        <br/>
                      </td>
                    </tr>
                    <tr>
                      <td>N Studis</td>
                      <td colspan="2">
                          <p>{tn}</p>
                      </td>
                    </tr>
                    <tr>
                      <td>Nachteilsausgleiche</td>
                      <td colspan="2">
                        <br/>
                      </td>
                    </tr>
                    <tr>
                      <td colspan="1">Nachmeldungen</td>
                      <td colspan="2">
                        <br/>
                      </td>
                    </tr>
                  </tbody>
                  <colgroup> <col style="width: 30.86%;"/> <col style="width: 5.75127%;"/> <col style="width: 63.4444%;"/> </colgroup>
                </table>
                <p class="auto-cursor-target">
                  <br/>
                </p>
              </ac:layout-cell>
              <ac:layout-cell>
                <h1 class="auto-cursor-target">Dokumente und Deadlines</h1>
                <table class="relative-table wrapped" style="width: 100.0%;">
                  <colgroup> <col style="width: 17.1451%;"/> <col style="width: 10.8261%;"/> <col style="width: 72.0846%;"/> </colgroup>
                  <tbody>
                    <tr>
                      <th colspan="3">Deadlines</th>
                    </tr>
                    <tr>
                      <td>Wordvorlage</td>
                      <td colspan="2">
                        <div class="content-wrapper">
                          <ac:task-list>
          <ac:task>
          <ac:task-id>18</ac:task-id>
          <ac:task-status>incomplete</ac:task-status>
          <ac:task-body>
                                <ac:link>
                                  <ri:user ri:userkey="20ad2a7e80924e4c018093417b7b0005"/>
                                </ac:link> Word-Vorlage bis zum <time datetime="{word_vorlage}"/> </ac:task-body>
          </ac:task>
          </ac:task-list>
                        </div>
                      </td>
                    </tr>
                    <tr>
                      <td>RM</td>
                      <td colspan="2">
                        <div class="content-wrapper">
                          <ac:task-list>
          <ac:task>
          <ac:task-id>20</ac:task-id>
          <ac:task-status>incomplete</ac:task-status>
          <ac:task-body>
                                <ac:link>
                                  <ri:user ri:userkey="20ad2a7e7dd62460017de19e99010007"/>
                                  <ac:plain-text-link-body><![CDATA[E-Examinations Rückmeldung]]></ac:plain-text-link-body>
                                </ac:link>  RM bis zum <time datetime="9999-01-01"/> </ac:task-body>
          </ac:task>
          </ac:task-list>
                        </div>
                      </td>
                    </tr>
                    <tr>
                      <td>TN</td>
                      <td colspan="2">
                        <div class="content-wrapper">
                          <ac:task-list>
          <ac:task>
          <ac:task-id>22</ac:task-id>
          <ac:task-status>incomplete</ac:task-status>
          <ac:task-body>
                                <ac:link>
                                  <ri:user ri:userkey="20ad2a7e80924e4c018093417b7b0005"/>
                                </ac:link> TN-Vorlage bis zum <time datetime="{tn_vorlage}"/> </ac:task-body>
          </ac:task>
          </ac:task-list>
                        </div>
                      </td>
                    </tr>
                  </tbody>
                </table>
                <p class="auto-cursor-target">
                  <br/>
                </p>
                <table class="relative-table wrapped">
                  <colgroup> <col style="width: 11.5423%;"/> <col style="width: 88.516%;"/> </colgroup>
                  <tbody>
                    <tr>
                      <th colspan="2">LPLUS</th>
                    </tr>
                    <tr>
                      <td>Lizenz</td>
                      <td>
                        <p>{lizenzame}</p>
                      </td>
                    </tr>
                    <tr>
                      <td>Katalog</td>
                      <td>
                        <p>{katalogname_nk}</p>
                      </td>
                    </tr>
                    <tr>
                      <td colspan="1">Randomisierung<br/>der Aufgaben</td>
                      <td colspan="1">
                        <br/>
                      </td>
                    </tr>
                    <tr>
                      <td colspan="1">Prüfungseinsicht</td>
                      <td colspan="1">
                        <div class="content-wrapper">
                          <p>
                            <br/>
                          </p>
                        </div>
                      </td>
                    </tr>
                  </tbody>
                </table>
                <p class="auto-cursor-target">
                  <br/>
                </p>
                <table class="relative-table wrapped" style="width: 100.0%;">
                  <colgroup> <col style="width: 17.1557%;"/> <col style="width: 21.2008%;"/> </colgroup>
                  <tbody>
                    <tr>
                      <th colspan="2">Anhänge</th>
                    </tr>
                    <tr>
                      <td>
                        <div class="content-wrapper">
                          <p>Wordvorlage</p>
                        </div>
                      </td>
                      <td>
                        <br/>
                      </td>
                    </tr>
                  </tbody>
                  <tbody>
                    <tr>
                      <td>TN-Liste</td>
                      <td>
                        <br/>
                      </td>
                    </tr>
                    <tr>
                      <td>Änderungswünsche</td>
                      <td>
                        <br/>
                      </td>
                    </tr>
                  </tbody>
                </table>
                <p>
                  <br/>
                </p>
                <p class="auto-cursor-target">
                  <br/>
                </p>
              </ac:layout-cell>
              <ac:layout-cell>
                <h1 class="auto-cursor-target">Dispatching</h1>
                <table class="relative-table wrapped" style="width: 100.0%;">
                  <colgroup> <col style="width: 27.827%;"/> <col style="width: 29.8963%;"/> <col style="width: 42.3119%;"/> </colgroup>
                  <tbody>
                    <tr>
                      <th>To Do</th>
                      <th>Wer?</th>
                      <th>E-Mail verschickt</th>
                    </tr>
                    <tr>
                      <td>
                        <div class="content-wrapper">
                          <p>Prüfungserstellung</p>
                        </div>
                      </td>
                      <td>
                        <div class="content-wrapper">
                          <ac:task-list>
          <ac:task>
          <ac:task-id>23</ac:task-id>
          <ac:task-status>incomplete</ac:task-status>
          <ac:task-body>
                                <span> </span>
                              </ac:task-body>
          </ac:task>
          </ac:task-list>
                        </div>
                      </td>
                      <td>
                        <div class="content-wrapper">
                          <ac:task-list>
          <ac:task>
          <ac:task-id>25</ac:task-id>
          <ac:task-status>incomplete</ac:task-status>
          <ac:task-body>
                                <span> </span>
                              </ac:task-body>
          </ac:task>
          </ac:task-list>
                        </div>
                      </td>
                    </tr>
                    <tr>
                      <td>QK</td>
                      <td>
                        <ac:task-list>
          <ac:task>
          <ac:task-id>92</ac:task-id>
          <ac:task-status>incomplete</ac:task-status>
          <ac:task-body> </ac:task-body>
          </ac:task>
          </ac:task-list>
                      </td>
                      <td>
                        <ac:task-list>
          <ac:task>
          <ac:task-id>28</ac:task-id>
          <ac:task-status>incomplete</ac:task-status>
          <ac:task-body>
                              <span> </span>
                            </ac:task-body>
          </ac:task>
          </ac:task-list>
                      </td>
                    </tr>
                    <tr>
                      <td>TN-Liste</td>
                      <td>
                        <ac:task-list>
          <ac:task>
          <ac:task-id>93</ac:task-id>
          <ac:task-status>incomplete</ac:task-status>
          <ac:task-body> </ac:task-body>
          </ac:task>
          </ac:task-list>
                      </td>
                      <td>
                        <ac:task-list>
          <ac:task>
          <ac:task-id>31</ac:task-id>
          <ac:task-status>incomplete</ac:task-status>
          <ac:task-body>
                              <span> </span>
                            </ac:task-body>
          </ac:task>
          </ac:task-list>
                      </td>
                    </tr>
                    <tr>
                      <td>Änderungswünsche</td>
                      <td>
                        <ac:task-list>
          <ac:task>
          <ac:task-id>94</ac:task-id>
          <ac:task-status>incomplete</ac:task-status>
          <ac:task-body> </ac:task-body>
          </ac:task>
          </ac:task-list>
                      </td>
                      <td>
                        <ac:task-list>
          <ac:task>
          <ac:task-id>34</ac:task-id>
          <ac:task-status>incomplete</ac:task-status>
          <ac:task-body>
                              <span> </span>
                            </ac:task-body>
          </ac:task>
          </ac:task-list>
                      </td>
                    </tr>
                    <tr>
                      <td colspan="3">
                        <div class="content-wrapper">
                          <ac:task-list>
          <ac:task>
          <ac:task-id>38</ac:task-id>
          <ac:task-status>incomplete</ac:task-status>
          <ac:task-body>Prüfung final abgenommen </ac:task-body>
          </ac:task>
          </ac:task-list>
                        </div>
                      </td>
                    </tr>
                  </tbody>
                </table>
                <p class="auto-cursor-target">
                  <br/>
                </p>
                <table class="relative-table wrapped">
                  <colgroup> <col style="width: 24.114%;"/> <col style="width: 75.9444%;"/> </colgroup>
                  <tbody>
                    <tr>
                      <th colspan="2">Besonderheiten und Anmerkungen</th>
                    </tr>
                    <tr>
                      <td>Prüfungserstellung</td>
                      <td>
                        <br/>
                      </td>
                    </tr>
                    <tr>
                      <td>QK</td>
                      <td>
                        <br/>
                      </td>
                    </tr>
                    <tr>
                      <td>TN-Liste</td>
                      <td>
                        <br/>
                      </td>
                    </tr>
                    <tr>
                      <td>Änderungswünsche</td>
                      <td>
                        <br/>
                      </td>
                    </tr>
                  </tbody>
                </table>
                <p class="auto-cursor-target">
                  <br/>
                </p>
                <table class="wrapped">
                  <colgroup> <col style="width: 404.0px;"/> <col style="width: 90.0px;"/> <col style="width: 29.0px;"/> </colgroup>
                  <tbody>
                    <tr>
                      <td class="highlight-#f4f5f7" colspan="1" data-highlight-colour="#f4f5f7" title="Hintergrundfarbe : Hellgrau 100 %">
                        <strong title="">Prüfungserstellung &amp; Qualitätskontrolle</strong>
                      </td>
                      <td class="highlight-#f4f5f7" colspan="1" data-highlight-colour="#f4f5f7" title="Hintergrundfarbe : Hellgrau 100 %">
                        <strong title="">Erstellung</strong>
                      </td>
                      <td class="highlight-#f4f5f7" colspan="1" data-highlight-colour="#f4f5f7" title="Hintergrundfarbe : Hellgrau 100 %">
                        <strong title="">OK</strong>
                      </td>
                    </tr>
                    <tr>
                      <td colspan="1">
                        <strong>Editor</strong>
                      </td>
                      <td>
                        <br/>
                      </td>
                      <td>
                        <br/>
                      </td>
                    </tr>
                    <tr>
                      <td colspan="1">
                        <p>(optional) Ordinal-Skalen</p>
                      </td>
                      <td>
                        <ac:task-list>
          <ac:task>
          <ac:task-id>95</ac:task-id>
          <ac:task-status>incomplete</ac:task-status>
          <ac:task-body> </ac:task-body>
          </ac:task>
          </ac:task-list>
                      </td>
                      <td>
                        <ac:task-list>
          <ac:task>
          <ac:task-id>96</ac:task-id>
          <ac:task-status>incomplete</ac:task-status>
          <ac:task-body> </ac:task-body>
          </ac:task>
          </ac:task-list>
                      </td>
                    </tr>
                    <tr>
                      <td colspan="1">
                        <p>(optional) wurde Layout übernommen und Punkte entfernt?</p>
                      </td>
                      <td>
                        <ac:task-list>
          <ac:task>
          <ac:task-id>97</ac:task-id>
          <ac:task-status>incomplete</ac:task-status>
          <ac:task-body> </ac:task-body>
          </ac:task>
          </ac:task-list>
                      </td>
                      <td>
                        <ac:task-list>
          <ac:task>
          <ac:task-id>98</ac:task-id>
          <ac:task-status>incomplete</ac:task-status>
          <ac:task-body> </ac:task-body>
          </ac:task>
          </ac:task-list>
                      </td>
                    </tr>
                    <tr>
                      <td colspan="1">
                        <p>Bepunktung kontrolliert</p>
                      </td>
                      <td colspan="1">
                        <ac:task-list>
          <ac:task>
          <ac:task-id>99</ac:task-id>
          <ac:task-status>incomplete</ac:task-status>
          <ac:task-body> </ac:task-body>
          </ac:task>
          </ac:task-list>
                      </td>
                      <td colspan="1">
                        <ac:task-list>
          <ac:task>
          <ac:task-id>100</ac:task-id>
          <ac:task-status>incomplete</ac:task-status>
          <ac:task-body> </ac:task-body>
          </ac:task>
          </ac:task-list>
                      </td>
                    </tr>
                    <tr>
                      <td colspan="1">
                        <p>Multiple-Choice Aufgaben überprüft</p>
                      </td>
                      <td colspan="1">
                        <ac:task-list>
          <ac:task>
          <ac:task-id>101</ac:task-id>
          <ac:task-status>incomplete</ac:task-status>
          <ac:task-body> </ac:task-body>
          </ac:task>
          </ac:task-list>
                      </td>
                      <td colspan="1">
                        <ac:task-list>
          <ac:task>
          <ac:task-id>102</ac:task-id>
          <ac:task-status>incomplete</ac:task-status>
          <ac:task-body> </ac:task-body>
          </ac:task>
          </ac:task-list>
                      </td>
                    </tr>
                    <tr>
                      <td colspan="1">
                        <p>Bepunktung kontrolliert</p>
                      </td>
                      <td colspan="1">
                        <ac:task-list>
          <ac:task>
          <ac:task-id>103</ac:task-id>
          <ac:task-status>incomplete</ac:task-status>
          <ac:task-body> </ac:task-body>
          </ac:task>
          </ac:task-list>
                      </td>
                      <td colspan="1">
                        <ac:task-list>
          <ac:task>
          <ac:task-id>104</ac:task-id>
          <ac:task-status>incomplete</ac:task-status>
          <ac:task-body> </ac:task-body>
          </ac:task>
          </ac:task-list>
                      </td>
                    </tr>
                    <tr>
                      <td colspan="1">
                        <p>Wiki-Link eingefügt</p>
                      </td>
                      <td colspan="1">
                        <ac:task-list>
          <ac:task>
          <ac:task-id>105</ac:task-id>
          <ac:task-status>incomplete</ac:task-status>
          <ac:task-body> </ac:task-body>
          </ac:task>
          </ac:task-list>
                      </td>
                      <td colspan="1">
                        <ac:task-list>
          <ac:task>
          <ac:task-id>106</ac:task-id>
          <ac:task-status>incomplete</ac:task-status>
          <ac:task-body> </ac:task-body>
          </ac:task>
          </ac:task-list>
                      </td>
                    </tr>
                    <tr>
                      <td colspan="1">
                        <p>Abschließende Kontrolle Editor/Plattform durchgeführt</p>
                      </td>
                      <td colspan="1">
                        <ac:task-list>
          <ac:task>
          <ac:task-id>107</ac:task-id>
          <ac:task-status>incomplete</ac:task-status>
          <ac:task-body> </ac:task-body>
          </ac:task>
          </ac:task-list>
                      </td>
                      <td colspan="1">
                        <ac:task-list>
          <ac:task>
          <ac:task-id>108</ac:task-id>
          <ac:task-status>incomplete</ac:task-status>
          <ac:task-body> </ac:task-body>
          </ac:task>
          </ac:task-list>
                      </td>
                    </tr>
                    <tr>
                      <td colspan="1">
                        <strong>Plattform</strong>
                      </td>
                      <td colspan="1">
                        <p>
                          <br/>
                        </p>
                      </td>
                      <td colspan="1">
                        <p>
                          <br/>
                        </p>
                      </td>
                    </tr>
                    <tr>
                      <td colspan="1">
                        <p>Richtige Zeit eingestellt</p>
                      </td>
                      <td colspan="1">
                        <ac:task-list>
          <ac:task>
          <ac:task-id>109</ac:task-id>
          <ac:task-status>incomplete</ac:task-status>
          <ac:task-body> </ac:task-body>
          </ac:task>
          </ac:task-list>
                      </td>
                      <td colspan="1">
                        <ac:task-list>
          <ac:task>
          <ac:task-id>110</ac:task-id>
          <ac:task-status>incomplete</ac:task-status>
          <ac:task-body> </ac:task-body>
          </ac:task>
          </ac:task-list>
                      </td>
                    </tr>
                    <tr>
                      <td colspan="1">Richtige Vorlagenauswahl getroffen</td>
                      <td colspan="1">
                        <ac:task-list>
          <ac:task>
          <ac:task-id>111</ac:task-id>
          <ac:task-status>incomplete</ac:task-status>
          <ac:task-body> </ac:task-body>
          </ac:task>
          </ac:task-list>
                      </td>
                      <td colspan="1">
                        <ac:task-list>
          <ac:task>
          <ac:task-id>112</ac:task-id>
          <ac:task-status>incomplete</ac:task-status>
          <ac:task-body> </ac:task-body>
          </ac:task>
          </ac:task-list>
                      </td>
                    </tr>
                    <tr>
                      <td colspan="1">Design-Optionen hinzugefügt</td>
                      <td colspan="1">
                        <ac:task-list>
          <ac:task>
          <ac:task-id>113</ac:task-id>
          <ac:task-status>incomplete</ac:task-status>
          <ac:task-body> </ac:task-body>
          </ac:task>
          </ac:task-list>
                      </td>
                      <td colspan="1">
                        <ac:task-list>
          <ac:task>
          <ac:task-id>114</ac:task-id>
          <ac:task-status>incomplete</ac:task-status>
          <ac:task-body> </ac:task-body>
          </ac:task>
          </ac:task-list>
                      </td>
                    </tr>
                    <tr>
                      <td colspan="1">Multiple-Choice-Einstellungen gesetzt</td>
                      <td colspan="1">
                        <ac:task-list>
          <ac:task>
          <ac:task-id>115</ac:task-id>
          <ac:task-status>incomplete</ac:task-status>
          <ac:task-body> </ac:task-body>
          </ac:task>
          </ac:task-list>
                      </td>
                      <td colspan="1">
                        <ac:task-list>
          <ac:task>
          <ac:task-id>116</ac:task-id>
          <ac:task-status>incomplete</ac:task-status>
          <ac:task-body> </ac:task-body>
          </ac:task>
          </ac:task-list>
                      </td>
                    </tr>
                    <tr>
                      <td colspan="1">Reportmappe erstellt</td>
                      <td colspan="1">
                        <ac:task-list>
          <ac:task>
          <ac:task-id>117</ac:task-id>
          <ac:task-status>incomplete</ac:task-status>
          <ac:task-body> </ac:task-body>
          </ac:task>
          </ac:task-list>
                      </td>
                      <td colspan="1">
                        <ac:task-list>
          <ac:task>
          <ac:task-id>118</ac:task-id>
          <ac:task-status>incomplete</ac:task-status>
          <ac:task-body> </ac:task-body>
          </ac:task>
          </ac:task-list>
                      </td>
                    </tr>
                    <tr>
                      <td colspan="1">Dozierende freigeschaltet</td>
                      <td colspan="1">
                        <ac:task-list>
          <ac:task>
          <ac:task-id>119</ac:task-id>
          <ac:task-status>incomplete</ac:task-status>
          <ac:task-body> </ac:task-body>
          </ac:task>
          </ac:task-list>
                      </td>
                      <td colspan="1">
                        <ac:task-list>
          <ac:task>
          <ac:task-id>120</ac:task-id>
          <ac:task-status>incomplete</ac:task-status>
          <ac:task-body> </ac:task-body>
          </ac:task>
          </ac:task-list>
                      </td>
                    </tr>
                    <tr>
                      <td colspan="1">Dozierende in Nachbewertung eingetragen</td>
                      <td colspan="1">
                        <ac:task-list>
          <ac:task>
          <ac:task-id>121</ac:task-id>
          <ac:task-status>incomplete</ac:task-status>
          <ac:task-body> </ac:task-body>
          </ac:task>
          </ac:task-list>
                      </td>
                      <td colspan="1">
                        <ac:task-list>
          <ac:task>
          <ac:task-id>122</ac:task-id>
          <ac:task-status>incomplete</ac:task-status>
          <ac:task-body> </ac:task-body>
          </ac:task>
          </ac:task-list>
                      </td>
                    </tr>
                    <tr>
                      <td colspan="1">Reihenfolge der Fächer kontrolliert</td>
                      <td colspan="1">
                        <ac:task-list>
          <ac:task>
          <ac:task-id>123</ac:task-id>
          <ac:task-status>incomplete</ac:task-status>
          <ac:task-body> </ac:task-body>
          </ac:task>
          </ac:task-list>
                      </td>
                      <td colspan="1">
                        <ac:task-list>
          <ac:task>
          <ac:task-id>124</ac:task-id>
          <ac:task-status>incomplete</ac:task-status>
          <ac:task-body> </ac:task-body>
          </ac:task>
          </ac:task-list>
                      </td>
                    </tr>
                  </tbody>
                </table>
                <p>
                  <br/>
                </p>
                <p>
                  <br/>
                </p>
              </ac:layout-cell>
            </ac:layout-section>
          </ac:layout>
        """

      ###Für Support-Projekt###
      else:
        unterseitenname=datetime.datetime.strptime(datum, "%d.%m.%Y").strftime("%Y-%m-%d")+" [S] "+name_prüfung+" NK "+lvnummer+f" ({dozent})"
        
        unterseite_blanko=f"""
            <ac:layout>
              <ac:layout-section ac:type="three_equal">
                <ac:layout-cell>
                  <h1 class="auto-cursor-target">Allgemeine Informationen zur Prüfung</h1>
                  <table class="wrapped relative-table" style="width: 100.0%;">
                    <tbody>
                      <tr>
                        <th colspan="3">allgemeines zur Prüfung</th>
                      </tr>
                      <tr>
                        <td>LV-Nr.</td>
                        <td colspan="2">
                          <div class="content-wrapper">
                            <p>{lvnummer}</p>
                          </div>
                        </td>
                      </tr>
                      <tr>
                        <td>Name der Prüfung</td>
                        <td colspan="2">
                          <div class="content-wrapper">
                            <p>{name_prüfung}</p>
                          </div>
                        </td>
                      </tr>
                      <tr>
                        <td>Prüfer:in</td>
                        <td colspan="2">
                          <p>{dozent}, {mail}</p>
                        </td>
                      </tr>
                      <tr>
                        <td>Datum + Uhrzeit</td>
                        <td colspan="2">
                          <div class="content-wrapper">
                            <p>{datum}</p>
                            <p>Durchgänge:</p>
                            <ul>
                              {durchgänge_einzeln}
                            </ul>
                          </div>
                        </td>
                      </tr>
                      <tr>
                        <td>Dauer</td>
                        <td colspan="2">
                          <div class="content-wrapper">
                            <p>{prüfungsdauer}</p>
                          </div>
                        </td>
                      </tr>
                      <tr>
                        <td colspan="1">Repository Seite</td>
                        <td colspan="2">
                          <br/>
                        </td>
                      </tr>
                      <tr>
                        <td>HK/NK</td>
                        <td colspan="2">
                          <p>
                            Link zur HK/NK:
                          </p>
                        </td>
                      </tr>
                      <tr>
                        <td>Ort der Prüfung</td>
                        <td colspan="2">
                          <div class="content-wrapper">
                            <p>{ort}</p>
                          </div>
                        </td>
                      </tr>
                      <tr>
                        <td>Safe Exam Browser</td>
                        <td colspan="2">
                          <div class="content-wrapper">
                            <p>
                              <span class="templateparameter">nein</span>
                            </p>
                          </div>
                        </td>
                      </tr>
                    </tbody>
                    <colgroup> <col style="width: 23.7903%;"/> <col style="width: 29.974%;"/> <col style="width: 46.2914%;"/> </colgroup>
                    <tbody>
                      <tr>
                        <th colspan="3">Informationen zur Durchführung</th>
                      </tr>
                      <tr>
                        <td>Webex-Raum</td>
                        <td colspan="2">
                          <div class="content-wrapper">
                            <p>
                              <br/>
                            </p>
                          </div>
                        </td>
                      </tr>
                      <tr>
                        <td>Uhrzeit zum Erscheinen</td>
                        <td colspan="2">30 Minuten vor der Prüfung</td>
                      </tr>
                      <tr>
                        <td>Besonderheiten</td>
                        <td colspan="2">
                          <div class="content-wrapper">
                            <p>
                              <br/>
                            </p>
                          </div>
                        </td>
                      </tr>
                    </tbody>
                  </table>
                  <p class="auto-cursor-target">
                    <br/>
                  </p>
                  <table class="relative-table wrapped" style="width: 100.0%;">
                    <tbody>
                      <tr>
                        <th colspan="3">Informationen für Studierende</th>
                      </tr>
                      <tr>
                        <td>Login</td>
                        <td colspan="2">
                          <br/>
                        </td>
                      </tr>
                      <tr>
                        <td>N Studis</td>
                        <td colspan="2">
                          <p>{tn}</p>
                        </td>
                      </tr>
                      <tr>
                        <td>Nachteilsausgleiche</td>
                        <td colspan="2">
                          <p>
                            <br/>
                          </p>
                        </td>
                      </tr>
                      <tr>
                        <td colspan="1">Nachmeldungen</td>
                        <td colspan="2">
                          <p>
                            <br/>
                          </p>
                        </td>
                      </tr>
                    </tbody>
                    <colgroup> <col style="width: 30.86%;"/> <col style="width: 5.75127%;"/> <col style="width: 63.4444%;"/> </colgroup>
                  </table>
                  <p class="auto-cursor-target">
                    <br/>
                  </p>
                </ac:layout-cell>
                <ac:layout-cell>
                  <h1 class="auto-cursor-target">Dokumente und Deadlines</h1>
                  <table class="relative-table wrapped" style="width: 100.0%;">
                    <colgroup> <col style="width: 32.5762%;"/> <col style="width: 3.29784%;"/> <col style="width: 64.0589%;"/> </colgroup>
                    <tbody>
                      <tr>
                        <th colspan="3">Deadlines</th>
                      </tr>
                      <tr>
                        <td>
                          <p>Erstellung (Reminder ggf. automatisiert)</p>
                        </td>
                        <td colspan="2">
                          <div class="content-wrapper">
                            <ac:task-list>
            <ac:task>
            <ac:task-id>18</ac:task-id>
            <ac:task-status>incomplete</ac:task-status>
            <ac:task-body>
                                  <ac:link>
                                    <ri:user ri:userkey="20ad2a7e82f723140182f8233a530008"/>
                                  </ac:link> Lehrende Erstellung Prüfung bis zum <time datetime="{prüfungserstellung}"/> </ac:task-body>
            </ac:task>
            </ac:task-list>
                          </div>
                        </td>
                      </tr>
                      <tr>
                        <td colspan="1">Plattformeinstellungen (Reminder ggf. automatisiert)</td>
                        <td colspan="2">
                          <div class="content-wrapper">
                            <ac:task-list>
            <ac:task>
            <ac:task-id>2199</ac:task-id>
            <ac:task-status>incomplete</ac:task-status>
            <ac:task-body>
                                  <span class="placeholder-inline-tasks"> <span class="fabric"> <span> <span> <ac:link>
                                            <ri:user ri:userkey="20ad2a7e82f723140182f824006c0009"/>
                                          </ac:link>  Plattformeinstellungen in LPLUS bis zum <time datetime="{plattformeinstellungen}"/> <br/>
                                        </span> </span> </span> </span>
                                </ac:task-body>
            </ac:task>
            </ac:task-list>
                          </div>
                        </td>
                      </tr>
                      <tr>
                        <td>Rückmeldung zu Änderungswünschen und Prüfungsabnahme von Lehrenden</td>
                        <td colspan="2">
                          <div class="content-wrapper">
                            <ac:task-list>
            <ac:task>
            <ac:task-id>2215</ac:task-id>
            <ac:task-status>incomplete</ac:task-status>
            <ac:task-body>
                                  <span class="placeholder-inline-tasks"> <ac:link>
                                      <ri:user ri:userkey="20ad2a7e82f723140182f82586ca000b"/>
                                      <ac:plain-text-link-body><![CDATA[Lehrende-Rückmeldung]]></ac:plain-text-link-body>
                                    </ac:link> RM bis zum <time datetime="9999-01-01"/> <br/>
                                  </span>
                                </ac:task-body>
            </ac:task>
            </ac:task-list>
                          </div>
                        </td>
                      </tr>
                      <tr>
                        <td>TN</td>
                        <td colspan="2">
                          <div class="content-wrapper">
                            <ac:task-list>
            <ac:task>
            <ac:task-id>22</ac:task-id>
            <ac:task-status>incomplete</ac:task-status>
            <ac:task-body>
                                  <ac:link>
                                    <ri:user ri:userkey="20ad2a7e80924e4c018093417b7b0005"/>
                                  </ac:link> <span style="color: rgb(23,43,77);">TN-Vorlage bis zum <time datetime="{tn_vorlage}"/> </span>
                                </ac:task-body>
            </ac:task>
            </ac:task-list>
                          </div>
                        </td>
                      </tr>
                    </tbody>
                  </table>
                  <p class="auto-cursor-target">
                    <br/>
                  </p>
                  <table class="relative-table wrapped">
                    <colgroup> <col style="width: 11.5423%;"/> <col style="width: 88.516%;"/> </colgroup>
                    <tbody>
                      <tr>
                        <th colspan="2">LPLUS</th>
                      </tr>
                      <tr>
                        <td>Lizenz</td>
                        <td>
                          <p>{lizenzame}</p>
                        </td>
                      </tr>
                      <tr>
                        <td>Katalog</td>
                        <td>
                          <p>
                            <p>{katalogname_nk}</p>
                          </p>
                        </td>
                      </tr>
                      <tr>
                        <td colspan="1">Repository</td>
                        <td colspan="1">
                          <p>
                            <br/>
                          </p>
                        </td>
                      </tr>
                      <tr>
                        <td colspan="1">Randomisierung<br/>der Aufgaben</td>
                        <td colspan="1">
                          <br/>
                        </td>
                      </tr>
                      <tr>
                        <td colspan="1">Prüfungseinsicht</td>
                        <td colspan="1">
                          <div class="content-wrapper">
                            <p>
                              <br/>
                            </p>
                          </div>
                        </td>
                      </tr>
                    </tbody>
                  </table>
                  <p class="auto-cursor-target">
                    <br/>
                  </p>
                  <table class="relative-table wrapped" style="width: 100.0%;">
                    <colgroup> <col style="width: 24.912%;"/> <col style="width: 28.2303%;"/> </colgroup>
                    <tbody>
                      <tr>
                        <th colspan="2">Anhänge</th>
                      </tr>
                      <tr>
                        <td>
                          <div class="content-wrapper">
                            <p>Plattformeinstellungen</p>
                          </div>
                        </td>
                        <td>
                          <div class="content-wrapper">
                            <p>
                              <br/>
                            </p>
                            <p>
                              <span class="placeholder-inline-tasks"> <br/>
                              </span>
                            </p>
                          </div>
                        </td>
                      </tr>
                      <tr>
                        <td>TN-Liste</td>
                        <td>
                          <div class="content-wrapper">
                            <p>
                              <br/>
                            </p>
                            <p>
                              <br/>
                            </p>
                          </div>
                        </td>
                      </tr>
                    </tbody>
                  </table>
                  <p>
                    <br/>
                  </p>
                  <p class="auto-cursor-target">
                    <br/>
                  </p>
                </ac:layout-cell>
                <ac:layout-cell>
                  <h1 class="auto-cursor-target">Dispatching</h1>
                  <table class="relative-table wrapped" style="width: 100.0%;">
                    <colgroup> <col style="width: 27.7526%;"/> <col style="width: 50.5254%;"/> <col style="width: 21.7195%;"/> </colgroup>
                    <tbody>
                      <tr>
                        <th>To Do</th>
                        <th>Wer?</th>
                        <th>E-Mail verschickt</th>
                      </tr>
                      <tr>
                        <td>QK Editor + Transfer 1</td>
                        <td>
                          <div class="content-wrapper">
                            <ac:task-list>
            <ac:task>
            <ac:task-id>2253</ac:task-id>
            <ac:task-status>incomplete</ac:task-status>
            <ac:task-body>
                                  <span> </span>
                                </ac:task-body>
            </ac:task>
            </ac:task-list>
                          </div>
                        </td>
                        <td>
                          <ac:task-list>
            <ac:task>
            <ac:task-id>28</ac:task-id>
            <ac:task-status>incomplete</ac:task-status>
            <ac:task-body>
                                <span> </span>
                              </ac:task-body>
            </ac:task>
            </ac:task-list>
                        </td>
                      </tr>
                      <tr>
                        <td>QK Editor + Transfer 2</td>
                        <td>
                          <ac:task-list>
            <ac:task>
            <ac:task-id>2256</ac:task-id>
            <ac:task-status>incomplete</ac:task-status>
            <ac:task-body>
                                <span> </span>
                              </ac:task-body>
            </ac:task>
            </ac:task-list>
                        </td>
                        <td>
                          <ac:task-list>
            <ac:task>
            <ac:task-id>31</ac:task-id>
            <ac:task-status>incomplete</ac:task-status>
            <ac:task-body>
                                <span> </span>
                              </ac:task-body>
            </ac:task>
            </ac:task-list>
                        </td>
                      </tr>
                      <tr>
                        <td>TN-Liste</td>
                        <td>
                          <ac:task-list>
            <ac:task>
            <ac:task-id>2257</ac:task-id>
            <ac:task-status>incomplete</ac:task-status>
            <ac:task-body>
                                <span> </span>
                              </ac:task-body>
            </ac:task>
            </ac:task-list>
                        </td>
                        <td>
                          <ac:task-list>
            <ac:task>
            <ac:task-id>34</ac:task-id>
            <ac:task-status>incomplete</ac:task-status>
            <ac:task-body>
                                <span> </span>
                              </ac:task-body>
            </ac:task>
            </ac:task-list>
                        </td>
                      </tr>
                      <tr>
                        <td colspan="3">
                          <div class="content-wrapper">
                            <ac:task-list>
            <ac:task>
            <ac:task-id>38</ac:task-id>
            <ac:task-status>incomplete</ac:task-status>
            <ac:task-body>Prüfung durch Lehrende final abgenommen</ac:task-body>
            </ac:task>
            </ac:task-list>
                          </div>
                        </td>
                      </tr>
                    </tbody>
                  </table>
                  <p class="auto-cursor-target">
                    <br/>
                  </p>
                  <table class="relative-table wrapped" style="width: 100.0%;">
                    <colgroup> <col style="width: 13.7495%;"/> <col style="width: 86.3089%;"/> </colgroup>
                    <tbody>
                      <tr>
                        <th colspan="2">Besonderheiten und Anmerkungen</th>
                      </tr>
                      <tr>
                        <td>Transfer 1</td>
                        <td>
                          <br/>
                        </td>
                      </tr>
                      <tr>
                        <td>Transfer 2</td>
                        <td>
                          <br/>
                        </td>
                      </tr>
                      <tr>
                        <td>TN-Liste</td>
                        <td>
                          <br/>
                        </td>
                      </tr>
                    </tbody>
                  </table>
                  <p class="auto-cursor-target">
                    <br/>
                  </p>
                  <table class="wrapped relative-table" style="width: 100.0%;">
                    <colgroup> <col style="width: 67.474%;"/> <col style="width: 16.2086%;"/> <col style="width: 16.3757%;"/> </colgroup>
                    <tbody>
                      <tr>
                        <th>Checkliste Editor QK</th>
                        <th colspan="1">Transfer 1</th>
                        <th colspan="1">Transfer 2</th>
                      </tr>
                      <tr>
                        <td>
                          <p>Schriftart prüfen (Aufgaben: Arial 12 fett, Antworten: Arial 12)</p>
                        </td>
                        <td colspan="1">
                          <ac:task-list>
            <ac:task>
            <ac:task-id>2234</ac:task-id>
            <ac:task-status>incomplete</ac:task-status>
            <ac:task-body>
                                <span> </span>
                              </ac:task-body>
            </ac:task>
            </ac:task-list>
                        </td>
                        <td colspan="1">
                          <ac:task-list>
            <ac:task>
            <ac:task-id>2235</ac:task-id>
            <ac:task-status>incomplete</ac:task-status>
            <ac:task-body>
                                <span> </span>
                              </ac:task-body>
            </ac:task>
            </ac:task-list>
                        </td>
                      </tr>
                      <tr>
                        <td>
                          <p>Hinweis-Button eingefügt → nicht bei Staatsexamina </p>
                        </td>
                        <td colspan="1">
                          <ac:task-list>
            <ac:task>
            <ac:task-id>2236</ac:task-id>
            <ac:task-status>incomplete</ac:task-status>
            <ac:task-body>
                                <span> </span>
                              </ac:task-body>
            </ac:task>
            </ac:task-list>
                        </td>
                        <td colspan="1">
                          <ac:task-list>
            <ac:task>
            <ac:task-id>2237</ac:task-id>
            <ac:task-status>incomplete</ac:task-status>
            <ac:task-body>
                                <span> </span>
                              </ac:task-body>
            </ac:task>
            </ac:task-list>
                        </td>
                      </tr>
                      <tr>
                        <td>
                          <p>Auflösung der Abbildungen überprüfen</p>
                        </td>
                        <td colspan="1">
                          <ac:task-list>
            <ac:task>
            <ac:task-id>2238</ac:task-id>
            <ac:task-status>incomplete</ac:task-status>
            <ac:task-body>
                                <span> </span>
                              </ac:task-body>
            </ac:task>
            </ac:task-list>
                        </td>
                        <td colspan="1">
                          <ac:task-list>
            <ac:task>
            <ac:task-id>2239</ac:task-id>
            <ac:task-status>incomplete</ac:task-status>
            <ac:task-body>
                                <span> </span>
                              </ac:task-body>
            </ac:task>
            </ac:task-list>
                        </td>
                      </tr>
                      <tr>
                        <td>
                          <p>Hinweistext auf Anlage im Aufgabetext ergänzt</p>
                        </td>
                        <td colspan="1">
                          <ac:task-list>
            <ac:task>
            <ac:task-id>2240</ac:task-id>
            <ac:task-status>incomplete</ac:task-status>
            <ac:task-body>
                                <span> </span>
                              </ac:task-body>
            </ac:task>
            </ac:task-list>
                        </td>
                        <td colspan="1">
                          <ac:task-list>
            <ac:task>
            <ac:task-id>2241</ac:task-id>
            <ac:task-status>incomplete</ac:task-status>
            <ac:task-body>
                                <span> </span>
                              </ac:task-body>
            </ac:task>
            </ac:task-list>
                        </td>
                      </tr>
                      <tr>
                        <td colspan="1">
                          <p>Layout prüfen (leere Zeilen, zu viel Platz zwischen Fragen und Antworten)</p>
                        </td>
                        <td colspan="1">
                          <ac:task-list>
            <ac:task>
            <ac:task-id>2242</ac:task-id>
            <ac:task-status>incomplete</ac:task-status>
            <ac:task-body>
                                <span> </span>
                              </ac:task-body>
            </ac:task>
            </ac:task-list>
                        </td>
                        <td colspan="1">
                          <ac:task-list>
            <ac:task>
            <ac:task-id>2243</ac:task-id>
            <ac:task-status>incomplete</ac:task-status>
            <ac:task-body>
                                <span> </span>
                              </ac:task-body>
            </ac:task>
            </ac:task-list>
                        </td>
                      </tr>
                      <tr>
                        <td colspan="1">
                          <p>Teilbewertung (keine Maluspunkte, richtige Teilbewertung ausgewählt)</p>
                        </td>
                        <td colspan="1">
                          <ac:task-list>
            <ac:task>
            <ac:task-id>2244</ac:task-id>
            <ac:task-status>incomplete</ac:task-status>
            <ac:task-body>
                                <span> </span>
                              </ac:task-body>
            </ac:task>
            </ac:task-list>
                        </td>
                        <td colspan="1">
                          <ac:task-list>
            <ac:task>
            <ac:task-id>2245</ac:task-id>
            <ac:task-status>incomplete</ac:task-status>
            <ac:task-body>
                                <span> </span>
                              </ac:task-body>
            </ac:task>
            </ac:task-list>
                        </td>
                      </tr>
                      <tr>
                        <td colspan="1">
                          <p>Aufgaben mit festem Format prüfen</p>
                        </td>
                        <td colspan="1">
                          <ac:task-list>
            <ac:task>
            <ac:task-id>2246</ac:task-id>
            <ac:task-status>incomplete</ac:task-status>
            <ac:task-body>
                                <span> </span>
                              </ac:task-body>
            </ac:task>
            </ac:task-list>
                        </td>
                        <td colspan="1">
                          <ac:task-list>
            <ac:task>
            <ac:task-id>2247</ac:task-id>
            <ac:task-status>incomplete</ac:task-status>
            <ac:task-body>
                                <span> </span>
                              </ac:task-body>
            </ac:task>
            </ac:task-list>
                        </td>
                      </tr>
                      <tr>
                        <td colspan="1">
                          <p>Dropdown und Lückentexte prüfen (Antwortoptionen und Bepunktung)</p>
                        </td>
                        <td colspan="1">
                          <ac:task-list>
            <ac:task>
            <ac:task-id>2248</ac:task-id>
            <ac:task-status>incomplete</ac:task-status>
            <ac:task-body>
                                <span> </span>
                              </ac:task-body>
            </ac:task>
            </ac:task-list>
                        </td>
                        <td colspan="1">
                          <ac:task-list>
            <ac:task>
            <ac:task-id>2249</ac:task-id>
            <ac:task-status>incomplete</ac:task-status>
            <ac:task-body>
                                <span> </span>
                              </ac:task-body>
            </ac:task>
            </ac:task-list>
                        </td>
                      </tr>
                      <tr>
                        <td colspan="1">
                          <p>Allgemeine Auffälligkeiten (z.B. sind die Aufgaben einheitlich erstellt?) - Bitte notieren/ Rücksprache halten</p>
                        </td>
                        <td colspan="1">
                          <ac:task-list>
            <ac:task>
            <ac:task-id>2250</ac:task-id>
            <ac:task-status>incomplete</ac:task-status>
            <ac:task-body>
                                <span> </span>
                              </ac:task-body>
            </ac:task>
            </ac:task-list>
                        </td>
                        <td colspan="1">
                          <ac:task-list>
            <ac:task>
            <ac:task-id>2251</ac:task-id>
            <ac:task-status>incomplete</ac:task-status>
            <ac:task-body>
                                <span> </span>
                              </ac:task-body>
            </ac:task>
            </ac:task-list>
                        </td>
                      </tr>
                    </tbody>
                  </table>
                  <p class="auto-cursor-target">
                    <br/>
                  </p>
                  <table class="relative-table wrapped" style="width: 100.0%;">
                    <colgroup> <col style="width: 83.3404%;"/> <col style="width: 16.718%;"/> </colgroup>
                    <tbody>
                      <tr>
                        <th colspan="1">Plattform Checkliste</th>
                        <th>QK</th>
                      </tr>
                      <tr>
                        <td colspan="1">
                          <p>Richtige Zeit eingestellt</p>
                        </td>
                        <td colspan="1">
                          <ac:task-list>
            <ac:task>
            <ac:task-id>2166</ac:task-id>
            <ac:task-status>incomplete</ac:task-status>
            <ac:task-body> </ac:task-body>
            </ac:task>
            </ac:task-list>
                        </td>
                      </tr>
                      <tr>
                        <td colspan="1">Richtige Vorlagenauswahl getroffen</td>
                        <td colspan="1">
                          <ac:task-list>
            <ac:task>
            <ac:task-id>2168</ac:task-id>
            <ac:task-status>incomplete</ac:task-status>
            <ac:task-body> </ac:task-body>
            </ac:task>
            </ac:task-list>
                        </td>
                      </tr>
                      <tr>
                        <td colspan="1">Design-Optionen hinzugefügt</td>
                        <td colspan="1">
                          <ac:task-list>
            <ac:task>
            <ac:task-id>2170</ac:task-id>
            <ac:task-status>incomplete</ac:task-status>
            <ac:task-body> </ac:task-body>
            </ac:task>
            </ac:task-list>
                        </td>
                      </tr>
                      <tr>
                        <td colspan="1">Multiple-Choice-Einstellungen gesetzt</td>
                        <td colspan="1">
                          <ac:task-list>
            <ac:task>
            <ac:task-id>2172</ac:task-id>
            <ac:task-status>incomplete</ac:task-status>
            <ac:task-body> </ac:task-body>
            </ac:task>
            </ac:task-list>
                        </td>
                      </tr>
                      <tr>
                        <td colspan="1">Reportmappe erstellt</td>
                        <td colspan="1">
                          <ac:task-list>
            <ac:task>
            <ac:task-id>2174</ac:task-id>
            <ac:task-status>incomplete</ac:task-status>
            <ac:task-body> </ac:task-body>
            </ac:task>
            </ac:task-list>
                        </td>
                      </tr>
                      <tr>
                        <td colspan="1">Dozierende freigeschaltet</td>
                        <td colspan="1">
                          <ac:task-list>
            <ac:task>
            <ac:task-id>2176</ac:task-id>
            <ac:task-status>incomplete</ac:task-status>
            <ac:task-body> </ac:task-body>
            </ac:task>
            </ac:task-list>
                        </td>
                      </tr>
                      <tr>
                        <td colspan="1">Dozierende + Sys-Admins in Nachbewertung eingetragen</td>
                        <td colspan="1">
                          <ac:task-list>
            <ac:task>
            <ac:task-id>2178</ac:task-id>
            <ac:task-status>incomplete</ac:task-status>
            <ac:task-body> </ac:task-body>
            </ac:task>
            </ac:task-list>
                        </td>
                      </tr>
                      <tr>
                        <td colspan="1">Reihenfolge der Fächer kontrolliert</td>
                        <td colspan="1">
                          <ac:task-list>
            <ac:task>
            <ac:task-id>2192</ac:task-id>
            <ac:task-status>incomplete</ac:task-status>
            <ac:task-body> </ac:task-body>
            </ac:task>
            </ac:task-list>
                        </td>
                      </tr>
                    </tbody>
                  </table>
                  <p class="auto-cursor-target">
                    <br/>
                  </p>
                </ac:layout-cell>
              </ac:layout-section>
            </ac:layout>

        
        """
      #Posten der HK-Seite im Wiki
      unterseite_nk_url=seite_post(unterseitenname,unterseite_blanko)
      gesamtübersicht[count]["Wiki-Unterseite NK"]=unterseite_nk_url
    else:
      unterseite_nk_url=""
    ###Links für HK/NK in die jeweiligen Seiten einfügen###
     
    def wiki_inhalt_abrufen(unterseite_hk_url):
        wiki_token=f"Bearer {auth_token}"
        #Inhalt manipulieren
        url = f'https://wikis.fu-berlin.de/rest/api/content/{unterseite_hk_url}?expand=body.storage'

        headers = {
        "Accept": "application/json;charset=UTF-8",
        "Content-Type": "application/json;charset=UTF-8",
        "Authorization": wiki_token
        }

        r = requests.get(url=url, headers=headers)

        ###Oberer Inhalt
        inhalt=json.loads(r.text)
        inhalt=inhalt["body"]["storage"]["value"]
        return inhalt

    ##Zuerst HK
    #Wiki-URL generieren + Inhalt abrufen
    if i["Hauptprüfung"][0]["Durchgänge"]!=None:
      unterseite_hk_url_kurz=unterseite_hk_url
      while "=" in unterseite_hk_url_kurz:
        unterseite_hk_url_kurz=unterseite_hk_url_kurz[1:]

      inhalt=wiki_inhalt_abrufen(unterseite_hk_url_kurz)
    
      unterseite_nk_url_kurz=unterseite_nk_url
      while "=" in unterseite_nk_url_kurz:
        unterseite_nk_url_kurz=unterseite_nk_url_kurz[1:]
        
      #Inhalt verändern
      inhalt=inhalt.replace(f'Link zur HK/NK:',f'Link zur HK/NK: <a href="{unterseite_nk_url}">{unterseite_nk_url}</a>')
    
    #Seite aktualisieren
    def wiki_inhalt_manipulieren(inhalt,unterseite_hk_url_kurz):
      wiki_token=f"Bearer {auth_token}"
      #Daten für das Seiten-Update ziehen
      url = f'https://wikis.fu-berlin.de/rest/api/content/{unterseite_hk_url_kurz}?expand=version'
      headers = {

      'Content-Type': 'application/json;charset=iso-8859-1',
      "Authorization": wiki_token
      }

      r = requests.get(url=url, headers=headers)
      seitenversion=r.json()["version"]

      seitenversion=seitenversion["_links"]["self"]

      while "/" in seitenversion:
          seitenversion=seitenversion[1:]

      seitentitel=r.json()["title"]

      url = f"https://wikis.fu-berlin.de/rest/api/content/{unterseite_hk_url_kurz}"

      headers = {
      "Accept": "application/json",
      "Content-Type": "application/json",
      "Authorization": wiki_token
      }

      payload = json.dumps( {
      "version": {
          "number":int(seitenversion)+1
      },
      "title": seitentitel,
      "type": "page",
      "status": "current",
      "ancestors": [],
      "body": {
          "storage": {
              "value": inhalt,
              "representation": "storage"
          }}
      } )


      response = requests.request(
      "PUT",
      url,
      data=payload,
      headers=headers
      )

      if response.status_code==200:
          print("Wiki-Seite erfolgreich aktualisiert")
    wiki_inhalt_manipulieren(inhalt,unterseite_hk_url_kurz)

    #Jetzt NK
    if i["Nachprüfung"][0]["Durchgänge"]!=None:
      unterseite_nk_url_kurz=unterseite_nk_url
      while "=" in unterseite_nk_url_kurz:
        unterseite_nk_url_kurz=unterseite_nk_url_kurz[1:]

      inhalt=wiki_inhalt_abrufen(unterseite_nk_url_kurz)
      
      #Inhalt verändern
      inhalt=inhalt.replace(f'Link zur HK/NK:',f'Link zur HK/NK: <a href="{unterseite_hk_url}">{unterseite_hk_url}</a>')
      
      #Seite aktualisieren
      wiki_inhalt_manipulieren(inhalt,unterseite_nk_url_kurz)


unterseite_generieren(gesamtübersicht)

def csv_mit_wiki_links_füttern():
    df = pandas.read_csv('Prüfungsverteilung_2111022_1649Uhr - Sortierung_Mailversand_abgeschlossen.csv',encoding = 'unicode_escape', sep=';') #Transformieren der CSV in pd, um Daten manipulieren zu können
    df = pandas.DataFrame(df)
    for index, row in df.iterrows():
      for eintrag in gesamtübersicht:
        if int(row["Index"])==int(eintrag["Index"]):
          df.at[index,"Wiki-Unterseite HK"] = eintrag["Wiki-Unterseite HK"]
          df.at[index,"Wiki-Unterseite NK"] = eintrag["Wiki-Unterseite NK"]
          row["Wiki-Unterseite NK"]=eintrag["Wiki-Unterseite NK"]

    df.to_csv('Prüfungsverteilung_2111022_1649Uhr - Sortierung_Mailversand_abgeschlossen.csv', index=False, sep=";", encoding='iso-8859-15')


csv_mit_wiki_links_füttern()