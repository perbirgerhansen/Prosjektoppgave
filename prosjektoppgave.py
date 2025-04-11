"""
Created on Wed Mar 12 10:25:24 2025



@author: Per birger Hansen
"""

import tkinter as tk
from tkinter import ttk
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import pandas as pd
import os
from datetime import datetime, timedelta
import math


#Kunde dictionary som hentes opp i kunde_combobox og kundeadr_textbox
kunde = {
    "Larvik Kommune": "Fayes gate 7, 3156 Larvik",
    "Færder kommune": "Tinghaugveien 16, 3163 Nøtterøy",
    "Horten Kommune": "Teatergata 11, 3187 Horten ",
    "Tønsberg kommune": "Halfdan Wilhelmsens alle 1, 3110 Tønsberg",
}


#hentes når skjema/forms lastes ved oppstart
def sisterekke():
    file_path = os.path.join('C:\\', 'prosjektoppgave', 'prosjektdata.xlsx')
    
    # Finnes filen og les den
    if os.path.exists(file_path):
        df = pd.read_excel(file_path)
        if not df.empty:
            last_row = df.iloc[-1]
            
             
            #Fyller inn teksten fra siste raden i Excel filen
            kunde_combobox.set(last_row["Kunde"])
            saksnr_textbox.config(state="normal")
            saksnr_textbox.delete("1.0", tk.END)
            saksnr_textbox.insert("1.0", str(last_row["Saksnr"]))
            saksnr_textbox.config(state="disabled")
            dato_textbox.delete("1.0", tk.END)
            dato_textbox.insert("1.0", last_row["Dato_Tid"])
            datostop_textbox.delete("1.0", tk.END)
            datostop_textbox.insert("1.0", last_row["Dato_stop"])
            tidbrukt_textbox.delete("1.0", tk.END)
            tidbrukt_textbox.insert("1.0", last_row["Tidbrukt"], "right")
            fornavn_textbox.delete("1.0", tk.END)
            fornavn_textbox.insert("1.0", last_row["Fornavn"])
            etternavn_textbox.delete("1.0", tk.END)
            etternavn_textbox.insert("1.0", last_row["Etternavn"])
            adresse_textbox.delete("1.0", tk.END)
            adresse_textbox.insert("1.0", last_row["Adresse"])
            telefon_textbox.delete(0, tk.END)
            telefon_textbox.insert(0, last_row["Telefon"])
            kategori_combobox.set(last_row["Kategori"])
            prioritet_combobox.set(last_row["Prioritet"])
            lisens_combobox.set(last_row["Lisens"])
            tilfredshet_combobox.set(last_row["Tilfredshet"])
            kostnad_textbox.delete("1.0", tk.END)
            kostnad_textbox.insert("1.0", last_row["Kostnad"], "right")
            beskrivelse_textbox.delete("1.0", tk.END)
            beskrivelse_textbox.insert("1.0", last_row["Beskrivelse"])    
            
           
            
        else:
            print("Excel-filen er tom.")
    else:
        print("Excel-filen finnes ikke.")




#Kakediagram
def priokake_graf(file_path):
           
    popup = tk.Toplevel()
    popup.title("Prioritetsdiagram")

    if os.path.exists(file_path):
        df = pd.read_excel(file_path)
        if not df.empty:
            
            # Antall prioriteringer
            priority_counts = df["Prioritet"].value_counts()
            
            # Lag diagram
            fig, ax = plt.subplots(figsize=(5, 5))
            priority_counts.plot(kind='pie', ax=ax, autopct='%1.1f%%', colors=['red', 'orange', 'green'], startangle=90)
            ax.set_title("Fordeling av Prioritet", fontsize=8)
            ax.set_ylabel("")  # Fjerner standard y-akse etikett
            ax.legend(priority_counts.index, loc="upper right", fontsize=6)

            # Legg diagrammet til popup-vinduet med FigureCanvasTkAgg
            canvas = FigureCanvasTkAgg(fig, master=popup)
            canvas_widget = canvas.get_tk_widget()
            canvas_widget.pack(fill="both", expand=True)
            canvas.draw()
        else:
            tk.Label(popup, text="Excel-filen er tom.").pack()
    else:
        tk.Label(popup, text="Excel-filen finnes ikke.").pack()

#Tilfredshet diagram 
def tilfredshet_graf(file_path):
    popup = tk.Toplevel()
    popup.title("Tilfredshet")

    if os.path.exists(file_path):
        df = pd.read_excel(file_path)
        if not df.empty:
            #Tilfredshetsskalaen
            def kategoriser_tilfredshet(verdi):
                if 1 <= verdi <= 6:
                    return "Lavt"
                elif 7 <= verdi <= 8:
                    return "Nøytralt"
                elif 9 <= verdi <= 10:
                    return "Positivt"

            df["Tilfredshet Gruppe"] = df["Tilfredshet"].apply(kategoriser_tilfredshet)

            #Regne ut prosent
            tilfreds_teller = df["Tilfredshet Gruppe"].value_counts(normalize=True) * 100

            #Diff mellom positive og negative
            prosent_positiv = tilfreds_teller.get("Positivt", 0)  # Henter verdi eller 0 hvis mangler
            prosent_negativ = tilfreds_teller.get("Lavt", 0)      # Henter verdi eller 0 hvis mangler
            differanse = prosent_positiv - prosent_negativ

            #Lag diagram
            fig, ax = plt.subplots(figsize=(6, 4))
            tilfreds_teller.plot(kind="bar", ax=ax, color=["red", "orange", "green"])
            ax.set_title(f"Fordeling av tilfredshet mellom Posetivt og negativt (%) - NPS: {differanse:.2f}%", fontsize=10)
            ax.set_xlabel("Gruppe", fontsize=8)
            ax.set_ylabel("Prosentandel", fontsize=8)
            ax.set_xticklabels(tilfreds_teller.index, rotation=0, fontsize=8)

            #Popup
            canvas = FigureCanvasTkAgg(fig, master=popup)
            canvas_widget = canvas.get_tk_widget()
            canvas_widget.pack(fill="both", expand=True)
            canvas.draw()
        else:
            tk.Label(popup, text="Excel-filen er tom.").pack()
    else:
        tk.Label(popup, text="Excel-filen finnes ikke.").pack()

#Henvendelse graf 
def henvendelser_graf(file_path):
    popup = tk.Toplevel()
    popup.title("Henvendelser etter tidsintervall")

    if os.path.exists(file_path):
        df = pd.read_excel(file_path)
        if not df.empty:
            
            df["Dato_Tid"] = pd.to_datetime(df["Dato_Tid"], errors='coerce')
            df = df.dropna(subset=["Dato_Tid"])  #Fjern rader med ugyldige Dato_tid

            #tidsintervaller
            def kategoriser_tidspunkt(tidspunkt):
                if 8 <= tidspunkt.hour < 10:
                    return "08-10"
                elif 10 <= tidspunkt.hour < 12:
                    return "10-12"
                elif 12 <= tidspunkt.hour < 14:
                    return "12-14"
                elif 14 <= tidspunkt.hour < 16:
                    return "14-16"
                else:
                    return "Utenfor intervall"

            df["Tidsintervall"] = df["Dato_Tid"].apply(kategoriser_tidspunkt)

            #antall henvendelser
            tidsintervall_teller = df["Tidsintervall"].value_counts()

            #Lag diagram
            fig, ax = plt.subplots(figsize=(6, 4))
            tidsintervall_teller.plot(kind="bar", ax=ax, color=["blue", "green", "orange", "red"])
            ax.set_title("Antall henvendelser per tidsintervall", fontsize=10)
            ax.set_xlabel("Tidsintervall", fontsize=8)
            ax.set_ylabel("Antall henvendelser", fontsize=8)
            ax.set_xticklabels(tidsintervall_teller.index, rotation=0, fontsize=8)

            
            canvas = FigureCanvasTkAgg(fig, master=popup)
            canvas_widget = canvas.get_tk_widget()
            canvas_widget.pack(fill="both", expand=True)
            canvas.draw()
        else:
            tk.Label(popup, text="Excel-filen er tom.").pack()
    else:
        tk.Label(popup, text="Excel-filen finnes ikke.").pack()
        
        
#Status melding i status feltet
def update_status_message(message):
    status_textbox.config(state="normal")
    status_textbox.delete("1.0", tk.END)
    status_textbox.insert("1.0", message)
    status_textbox.config(state="disabled")

#Øverfører til excel "Lage knap"
def skrive_til_kollone():
    
    
    #Hent Saksnr fra saksnr_textbox
    saksnr = saksnr_textbox.get("1.0", tk.END).strip()
    
    #Marker raden i treeview basert på Saksnr
    for item in treeview.get_children():
        values = treeview.item(item)["values"]
        if values and str(values[0]) == saksnr:
            treeview.selection_set(item)  
            treeview.focus(item)  
            break
    
    
    
    dato = dato_textbox.get("1.0", tk.END).strip()
    datostop = datostop_textbox.get("1.0", tk.END).strip()
    tidbrukt = tidbrukt_textbox.get("1.0", tk.END).strip()
    fornavn = fornavn_textbox.get("1.0", tk.END).strip()
    etternavn = etternavn_textbox.get("1.0", tk.END).strip()
    adresse = adresse_textbox.get("1.0", tk.END).strip()
    telefon = telefon_textbox.get().strip()
    kategori = kategori_combobox.get().strip()
    prioritet = prioritet_combobox.get().strip()
    lisens = lisens_combobox.get().strip()
    tilfredshet =  tilfredshet_combobox.get().strip()
    kostnad = kostnad_textbox.get("1.0", tk.END).strip()
    beskrivelse = beskrivelse_textbox.get("1.0", tk.END).strip()
    kunde = kunde_combobox.get().strip()
    
    
    if kunde and dato and datostop and tidbrukt and fornavn and etternavn and adresse and telefon and kategori and prioritet and lisens and  tilfredshet and kostnad and beskrivelse:
        file_path = os.path.join('C:\\', 'prosjektoppgave', 'prosjektdata.xlsx')

        #Sjekk om filen finnes
        if os.path.exists(file_path):
            df = pd.read_excel(file_path)
        else:
            df = pd.DataFrame(columns=["Saksnr", "Dato_Tid", "Dato_stop", "Tidbrukt", "Fornavn", "Etternavn", "Adresse", "Telefon", "Kategori", "Prioritet", "Lisens", "Henvendeleser", "kostnad", "Beskrivelse", "Kunde"])

        #Hent den markerte radens Saksnr fra Treeview
        selected_item = treeview.focus()

        if selected_item:
            selected_saksnr = treeview.item(selected_item)["values"][0]  # Hent Saksnr fra Treeview
           
          
            df.loc[df["Saksnr"] == selected_saksnr, :] = [int(selected_saksnr), str(dato), str(datostop), float(tidbrukt), str(fornavn), str(etternavn), str(adresse), int(telefon), str(kategori), str(prioritet), int(lisens), float( tilfredshet), float(kostnad), str(beskrivelse), str(kunde)]       
        
            
        else:
            saksnr = len(df) + 1
            new_row = {"Saksnr": saksnr, "Dato_Tid": dato, "Dato_stop": datostop, "Tidbrukt": tidbrukt, "Fornavn": fornavn, "Etternavn": etternavn, "Adresse": adresse, "Telefon": telefon, "Kategori": kategori, "Prioritet": prioritet, "Lisens": lisens, "Tilfredshet":  tilfredshet, "Kostnad": kostnad, "Beskrivelse": beskrivelse, "Kunde": kunde}
            df = pd.concat([df, pd.DataFrame([new_row])], ignore_index=True)

        #Lagre data til Excel
        df.to_excel(file_path, index=False)
        
        
        
        #Oppdater tabellen
        oppdater_tabell(file_path)

        update_status_message("Data lagret!")
    else:
        update_status_message("Vennligst fyll ut alle feltene før du lagrer.")


#Setter dato og klokke i dato textboksen brukes på start_button 
def settdato():
    dato_textbox.config(state="normal")
    current_date = datetime.now().strftime('%d-%m-%Y %H:%M:%S')
    dato_textbox.delete("1.0", tk.END)  
    dato_textbox.insert("1.0", current_date)
    dato_textbox.config(state="disabled")

#Setter stopdato og klokke i datostop textboksen og benyttes i settstop_brukstid for å regne ut antall minutter    
def settdatostop():
    current_date = datetime.now().strftime('%d-%m-%Y %H:%M:%S')
    datostop_textbox.delete("1.0", tk.END)  # Tømmer tekstfeltet
    datostop_textbox.insert("1.0", current_date)  # Setter inn dato og klokkeslett

#Brukes til å sette stopdato og klokke, funksjonen brukes på stop_button, henter også funksjon beregn_kostnad_fra_tid for å regne ut kostnad     
def settstop_brukttid():
    settdatostop()
    tidsdiff()
    beregn_kostnad_fra_tid()

def beregn_gjennomsnitt_tidbruk(file_path, tekstut_textbox):
    tekstut_textbox.delete("1.0", tk.END)
    try:
        #Leser filen
        df = pd.read_excel(file_path)

        #finnes kollonen
        if "Tidbrukt" in df.columns:
            #summer
            total_tidbruk = df["Tidbrukt"].sum()
            antall_rader = len(df["Tidbrukt"])
            gj_samtalletid= total_tidbruk / antall_rader if antall_rader > 0 else 0
            
            gj_samtalletid_rounded = math.ceil(gj_samtalletid)
            
            tekstut_textbox.insert("1.0", f"Gj. samtaletid: {gj_samtalletid_rounded} minutter")
        else:
            tekstut_textbox.insert("1.0", "Kolonnen 'Tidbrukt' mangler.")
    except Exception as e:
        tekstut_textbox.insert("1.0", f"Feil ved behandling: {e}")

#Benyttes til å lage ny rad/ny registrering, knyttet til ny_button 
def lagnyrad():
    tekstut_textbox.delete("1.0", tk.END)
    file_path = os.path.join('C:\\', 'prosjektoppgave', 'prosjektdata.xlsx')
 
    #Sjekk om filen finnes for å finne det neste saksnummeret
    if os.path.exists(file_path):
        df = pd.read_excel(file_path)
        next_saksnr = len(df) + 1 if not df.empty else 1
    else:
        next_saksnr = 1  
 
   
    #Oppdater Saksnr tekstboks med neste ledige saksnummer
    saksnr_textbox.config(state="normal")
    saksnr_textbox.delete("1.0", tk.END)
    saksnr_textbox.insert("1.0", str(next_saksnr))
    saksnr_textbox.config(state="disabled")
 
    #Hent dagens dato
    current_date = datetime.now().strftime('%d-%m-%Y %H:%M:%S')  
 
    #Sett datoen i tekstboksen
    dato_textbox.config(state="normal")  
    dato_textbox.delete("1.0", tk.END)
    dato_textbox.insert("1.0", current_date)
    dato_textbox.config(state="disabled")  
    datostop_textbox.config(state="normal")  
    datostop_textbox.delete("1.0", tk.END)
    tidbrukt_textbox.delete("1.0", tk.END)   
 
    #Tøm de andre tekstboksene
    fornavn_textbox.delete("1.0", tk.END)
    etternavn_textbox.delete("1.0", tk.END)
    adresse_textbox.delete("1.0", tk.END)
    telefon_textbox.delete(0, tk.END)
    kategori_combobox.set("")
    prioritet_combobox.set("")
    lisens_combobox.set("")
    tilfredshet_combobox.set("")
    kostnad_textbox.delete("1.0", tk.END)
    beskrivelse_textbox.delete("1.0", tk.END)
    status_textbox.config(state="normal")  
    status_textbox.delete("1.0", tk.END)   
    status_textbox.config(state="disabled")  
    kunde_combobox.set("")
    

#Regner ut kostnad antall minutter multipliseres med hundre kroner    
def beregn_kostnad_fra_tid():
    try:
        minutter = int(tidbrukt_textbox.get("1.0", tk.END).strip())  # Fjern eventuell ekstra blanke eller nye linjer
        kostnad = minutter * 100  #Multipliser med 100 kroner
        if kostnad == 0:
            kostnad = 1
        kostnad_textbox.delete("1.0", tk.END)
        kostnad_textbox.insert("1.0", f"{kostnad}")
    except ValueError:
        kostnad_textbox.delete("1.0", tk.END)
        kostnad_textbox.insert("1.0", "Ugyldig input")

def finn_hoyeste_og_laveste(data):
    hoyeste_verdi = data['Tidbrukt'].max()
    laveste_verdi = data['Tidbrukt'].min()
    return f"Lengste samtale: {hoyeste_verdi}\nKorteste samtale: {laveste_verdi}"

def kjor_hoy_lav():
    try:
        data = pd.read_excel(file_path)
        result = finn_hoyeste_og_laveste(data)
        tekstut_textbox.delete(1.0, tk.END)  # Fjern gammel tekst
        tekstut_textbox.insert(tk.END, result)  # Sett inn ny tekst
    except Exception as e:
        tekstut_textbox.delete(1.0, tk.END)
        tekstut_textbox.insert(tk.END, f"Feil: {str(e)}")

# Filsti
file_path = r"C:\prosjektoppgave\prosjektdata.xlsx"    
    
    
#Kjører diagram er knyttet til rapport knapp
def kjordiagram():
    file_path = r"C:\prosjektoppgave\prosjektdata.xlsx"
    priokake_graf(file_path)
    
def kjordiagram_tilfredshet():
    file_path = r"C:\prosjektoppgave\prosjektdata.xlsx"
    tilfredshet_graf(file_path)
    

def kjordiagram_henvendelser():
    file_path = r"C:\prosjektoppgave\prosjektdata.xlsx"
    henvendelser_graf(file_path)
    
def kjor_tidbruk_gj():
    file_path = r"C:\prosjektoppgave\prosjektdata.xlsx"
    beregn_gjennomsnitt_tidbruk(file_path, tekstut_textbox)

#Tømmer treeview og leser inn fra excel filen og legger inn i treeview 
def oppdater_tabell(file_path):
    for row in treeview.get_children():
        treeview.delete(row)
    
    if os.path.exists(file_path):
        df = pd.read_excel(file_path)
        for _, row in df.iterrows():
            treeview.insert("", "end", values=row.tolist())
            
def large_knap():
    tekstut_textbox.delete("1.0", tk.END)
    skrive_til_kollone()
    #tekstut_textbox.delete("1.0", tk.END)
    #file_path = os.path.join('C:\\', 'prosjektoppgave', 'prosjektdata.xlsx')
    #oppdater_tabell(file_path)
    

#Henter adresse som ligger i kunde  fra kunde_combobox og henter             
def vis_adresse(event):
    valgt_kunde = kunde_combobox.get()
    adresse = kunde.get(valgt_kunde, "") 
    kundeadr_textbox.config(state="normal") 
    kundeadr_textbox.delete("1.0", tk.END) 
    kundeadr_textbox.insert("1.0", adresse)
    kundeadr_textbox.config(state="disabled")
            

#Opprett hovedvinduet
root = tk.Tk()
root.title("Prosjekt oppgave i PY1010 2025 Per Hansen")
root.configure(bg='lightgray')
root.focus_force()

#Opprett hoved/main ramme
main_frame = tk.Frame(root)
main_frame.pack(padx=12, pady=12, fill="both", expand=True)


#setter regler for telefonnr.
def tlf_regler(action, value_if_allowed):
    if action == "1":  
        return value_if_allowed.isdigit() and len(value_if_allowed) <= 8
    return True

 
vcmd = (root.register(tlf_regler), "%d", "%P")


#Regner tidsdifferansen, og setter minimum verdi til 1 minutt ved avrunning  
def tidsdiff():
    try:
        #Hent datoer fra tekstboksene
        start_date_str = dato_textbox.get("1.0", tk.END).strip()
        stop_date_str = datostop_textbox.get("1.0", tk.END).strip()

        #Setter formate på dato og tid 
        start_date = datetime.strptime(start_date_str, "%d-%m-%Y %H:%M:%S")
        stop_date = datetime.strptime(stop_date_str, "%d-%m-%Y %H:%M:%S")

        #Regner ut differansen
        time_difference = stop_date - start_date
        if time_difference.total_seconds() <= 0:
            time_difference = timedelta(seconds=60) 
            
        
        total_minutes = math.ceil(time_difference.total_seconds() / 60)  
        
        
        
        #Viser tidsdifferansen i tekstboksen
        tidbrukt_textbox.delete("1.0", tk.END)
        tidbrukt_textbox.insert(tk.END, f"{total_minutes}")
    except ValueError:
        tidbrukt_textbox.delete("1.0", tk.END)
        tidbrukt_textbox.insert(tk.END, "Ugyldige datoer. Bruk formatet: YYYY-MM-DD HH:MM:SS")



#Oppretter kunde, combobox og et textfelt som henter verdier ut fra kunde dictionary
kunde_frame = tk.Frame(main_frame)
kunde_frame.pack(fill="x", pady=5)
label = tk.Label(kunde_frame, text="Velg kunde:   ")
label.pack(side="left")
kunde_combobox = ttk.Combobox(kunde_frame, values=list(kunde.keys()), state="readonly")
kunde_combobox.bind("<<ComboboxSelected>>", vis_adresse) 
kunde_combobox.pack(side="left")
kundeadr_textbox = tk.Text(kunde_frame, width=45, height=1, bg="#f0f0f0", highlightthickness=0, borderwidth=0, state="disabled")  # Deaktivert som standard
kundeadr_textbox.pack(side="left")



#Oppretter saksnr,startdato, stopdato, tidsbruk textboxer, pluss knapper start og stop  og frame
saksnr_dato_frame = tk.Frame(main_frame)
saksnr_dato_frame.pack(fill="x", pady=5)
tk.Label(saksnr_dato_frame, text="  Saksnr:", width=10).pack(side="left", padx=1, pady=0)
saksnr_textbox = tk.Text(saksnr_dato_frame, height=1, width=30)
saksnr_textbox.pack(side="left")
saksnr_textbox.insert("1.0", "1")
saksnr_textbox.config(state="disabled")
start_button = tk.Button(saksnr_dato_frame, text="Start ->", width=8, command=settdato)
start_button.pack(side="left", padx=5)
dato_textbox = tk.Text(saksnr_dato_frame, height=1, width=30)
dato_textbox.pack(side="left")
datostop_textbox = tk.Text(saksnr_dato_frame, height=1, width=30)
datostop_textbox.pack(side="left")
stop_button = tk.Button(saksnr_dato_frame, text="<-Stop", width=8, command=settstop_brukttid)
stop_button.pack(side="left", padx=5)
tidbrukt_textbox = tk.Text(saksnr_dato_frame, height=1, width=5,  bg="#f0f0f0", highlightthickness=0, borderwidth=0)
tidbrukt_textbox.pack(side="left")
tk.Label(saksnr_dato_frame, text=":Antall minutter * kr 100 = ", height=1).pack(side="left", padx=1, pady=0)


#opprett Etternavn og Fornavn textbox pluss frame
fornavn_etternavn_frame = tk.Frame(main_frame)
fornavn_etternavn_frame.pack(fill="x", pady=5)
tk.Label(fornavn_etternavn_frame, text="Fornavn:", width=10).pack(side="left", padx=1, pady=0)
fornavn_textbox = tk.Text(fornavn_etternavn_frame, height=1, width=30)
fornavn_textbox.pack(side="left")
tk.Label(fornavn_etternavn_frame, text="Etternavn:", width=10).pack(side="left", padx=1, pady=0)
etternavn_textbox = tk.Text(fornavn_etternavn_frame, height=1, width=30)
etternavn_textbox.pack(side="left")


#Oppretter adresse og telefon text box pluss frame
adresse_telefon_frame = tk.Frame(main_frame)
adresse_telefon_frame.pack(fill="x", pady=5)
tk.Label(adresse_telefon_frame, text="  Adresse:", width=10).pack(side="left", padx=1, pady=0)
adresse_textbox = tk.Text(adresse_telefon_frame, height=1, width=30)
adresse_textbox.pack(side="left")
tk.Label(adresse_telefon_frame, text="Telefon:", width=10).pack(side="left", padx=1, pady=0)
telefon_textbox = tk.Entry(adresse_telefon_frame, validate="key", validatecommand=vcmd, width=30)
telefon_textbox.pack(side="left")

#Oppretter kategori og prioritet combobox pluss frame
kategori_prio_frame = tk.Frame(main_frame)
kategori_prio_frame.pack(fill="x", pady=5)
tk.Label(kategori_prio_frame, text="Kategori:", width=10).pack(side="left", padx=1, pady=0)
kategori_combobox = ttk.Combobox(kategori_prio_frame, values=["Bruker", "Maskinvare", "Print", "Programvare", "Bestilling"], width=37)
kategori_combobox.pack(side="left")
kategori_combobox.set("")
tk.Label(kategori_prio_frame, text="Prioritet:", width=10).pack(side="left", padx=1, pady=0)
prioritet_combobox = ttk.Combobox(kategori_prio_frame, values=["Kritisk", "Normal", "Lav"], width=28)
prioritet_combobox.pack(side="left")
prioritet_combobox.set("")

#Oppretter Lisens og tilfredshet combobox pluss fram
lisens_tilfredshet_kost_frame = tk.Frame(main_frame)
lisens_tilfredshet_kost_frame.pack(fill="x", pady=5)
tk.Label(lisens_tilfredshet_kost_frame, text="Lisens kr.:", width=10).pack(side="left", padx=1, pady=0)
lisens_combobox = ttk.Combobox(lisens_tilfredshet_kost_frame, values=["1000", "2000", "3000"], width=37)
lisens_combobox.pack(side="left")
lisens_combobox.set("")
tk.Label(lisens_tilfredshet_kost_frame, text="Tilfredshet:", width=10).pack(side="left", padx=1, pady=0)
tilfredshet_combobox = ttk.Combobox(lisens_tilfredshet_kost_frame, values=["1", "2", "3", "4", "5", "6", "7", "8", "9", "10"], width=28)
tilfredshet_combobox.pack(side="left")
tilfredshet_combobox.set("")

#Opprett kostnad text box
kostnad_textbox = tk.Text(saksnr_dato_frame, height=1, width=6, bg="#f0f0f0", highlightthickness=0, borderwidth=0)
kostnad_textbox.tag_configure("right", justify="right")
kostnad_textbox.pack(side="left")
tk.Label(saksnr_dato_frame, text="Kroner", width=10).pack(side="left", padx=1, pady=0)

#Oppretter Beskrivelse textbox pluss frame 
beskrivelse_frame = tk.Frame(main_frame)
beskrivelse_frame.pack(fill="x", pady=15)
tk.Label(beskrivelse_frame, text="Beskrivelse:", width=10).pack(side="left", padx=1, pady=0)
beskrivelse_textbox = tk.Text(beskrivelse_frame, height=14, width=30)
beskrivelse_textbox.pack(side="left")

#Opprett knapper med frame
button_frame = tk.Frame(main_frame)
button_frame.pack(fill="x", pady=5)
lagre_button = tk.Button(button_frame, text="Lagre", width=13, command= large_knap)
lagre_button.pack(side="left", padx=5)
ny_button = tk.Button(button_frame, text="Ny", width=13, command=lagnyrad)
ny_button.pack(side="left", padx=5)
prioritet_button = tk.Button(button_frame, text="Prioritet", width=13, command=kjordiagram)
prioritet_button.pack(side="left", padx=5)
tilfredshet_button = tk.Button(button_frame, text="Tilfredshet-NPS", width=13, command=kjordiagram_tilfredshet)
tilfredshet_button.pack(side="left", padx=5)
henvendelser_button = tk.Button(button_frame, text="Henvendelser", width=13, command=kjordiagram_henvendelser)
henvendelser_button.pack(side="left", padx=5)
gjennomsnitt_button = tk.Button(button_frame, text="Gj. samtale", width=13, command=kjor_tidbruk_gj)
gjennomsnitt_button.pack(side="left", padx=5)
samtale_button = tk.Button(button_frame, text="Samtale høy/lav", width=13, command=kjor_hoy_lav)
samtale_button.pack(side="left", padx=5)
tekstut_textbox = tk.Text(button_frame, height=2, width=80, bg="#f0f0f0", highlightthickness=0, borderwidth=0)
tekstut_textbox.pack(side="left")

#Opprett statusbox med frame
status_frame = tk.Frame(main_frame)
status_frame.pack(fill="x", pady=15)
status_textbox = tk.Text(status_frame, height=1, width=45, bg="#f0f0f0", highlightthickness=0, borderwidth=0)
status_textbox.tag_configure("center", justify="center")
status_textbox.pack(side="left")

#Opprett en Treeview for å vise Excel-innhold inkl. scrollbar
treeview_frame = tk.Frame(main_frame)
treeview_frame.pack(fill="both", padx=10, pady=20)

scrollbar = ttk.Scrollbar(treeview_frame, orient=tk.VERTICAL)


#Definer kolonner for Treeview
columns = ("Saksnr", "Dato_tid", "Dato_stop", "Tidbrukt", "Fornavn", "Etternavn", "Adresse", "Telefon", 
           "Kategori", "Prioritet", "Lisens", "Tilfredshet","Kostnad", "Beskrivelse", "Kunde")


    
#Opprett Treeview med scrollbar
treeview = ttk.Treeview(beskrivelse_frame, columns=columns, show="headings", yscrollcommand=scrollbar.set)

scrollbar = ttk.Scrollbar(beskrivelse_frame, orient=tk.VERTICAL, command=treeview.yview)

#Konfigurer scrollbar til å styre Treeview
scrollbar.config(command=treeview.yview)

#Plasser scrollbar og Treeview i rammen
scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
treeview.pack(side=tk.LEFT, fill="both", padx=5, pady=5, expand=True)

#Legg til kolonner til Treeview
for col in columns:
    treeview.heading(col, text=col)
    treeview.column(col, width=40)

#Skjul kolonner i Treeview
hidden_columns = ["Kostnad", "Beskrivelse", "Tidbrukt", "Adresse", "Kategori", "Telefon", "Lisens", "Tilfredshet", "Prioritet" ]

#Legg til kolonner til Treeview
for col in columns:
    treeview.heading(col, text=col)
    if col in hidden_columns:
        treeview.column(col, width=0, stretch=False)  
    else:
        treeview.column(col, width=40)  


#Oppdater tabellen med data
file_path = os.path.join('C:\\', 'prosjektoppgave', 'prosjektdata.xlsx')
oppdater_tabell(file_path)

#Funksjon for å håndtere valg i Treeview
def oppdater_felter(event):
    selected_item = treeview.focus()
    if selected_item:
        values = treeview.item(selected_item, "values")
        if values:  
            
            
            #Oppdater tekstfeltene 
            saksnr_textbox.config(state="normal")
            saksnr_textbox.delete("1.0", tk.END)
            saksnr_textbox.insert("1.0", values[0])
            saksnr_textbox.config(state="disabled")

            dato_textbox.config(state="normal")
            dato_textbox.delete("1.0", tk.END)
            dato_textbox.insert("1.0", values[1])
            dato_textbox.config(state="disabled")

            datostop_textbox.delete("1.0", tk.END)
            datostop_textbox.insert("1.0", values[2])

            tidbrukt_textbox.delete("1.0", tk.END)
            tidbrukt_textbox.insert("1.0", values[3])

            fornavn_textbox.delete("1.0", tk.END)
            fornavn_textbox.insert("1.0", values[4])

            etternavn_textbox.delete("1.0", tk.END)
            etternavn_textbox.insert("1.0", values[5])

            adresse_textbox.delete("1.0", tk.END)
            adresse_textbox.insert("1.0", values[6])

            telefon_textbox.delete(0, tk.END)
            telefon_textbox.insert(0, values[7])

            kategori_combobox.set(values[8])
            prioritet_combobox.set(values[9])
            lisens_combobox.set(values[10])

            tilfredshet_combobox.set(values[11])
                        
            kostnad_textbox.delete("1.0", tk.END)
            kostnad_textbox.insert("1.0", values[12])

            beskrivelse_textbox.delete("1.0", tk.END)
            beskrivelse_textbox.insert("1.0", values[13])
            
            kunde_combobox.set(values[14])
            
            

#Oppdaterer Treeview
treeview.bind("<<TreeviewSelect>>", oppdater_felter)





#Kjøres ved oppstart
sisterekke()
root.mainloop()        
        




