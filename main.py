from ast import increment_lineno
from asyncore import read
import zipfile
import os
from bs4 import BeautifulSoup
from docx import Document
from docx.opc.constants import RELATIONSHIP_TYPE as RT
import json



FILES_NAMES = []
FILES_NAMES_JSON = []
FILES_NAMES_DOCX = []

HYPERLINK_MERGE = []
HYPERLINK_LINK = []
HYPERLINK_TEXT = []

# relation des fichiers (avec les liens externes)
FILES_RELATIONS = []
FILES_WEIGHT = {}

import mysql.connector
from mysql.connector import Error

import networkx as nx
import matplotlib.pyplot as plt

DATABASE_HOST = "localhost"
DATABASE_USER = "root"
DATABASE_PASSWORD = ""
DATABASE = "m1_info"

def file_to_xml(file_name) :

    try :

        file_docx = file_name+".docx"

        # unzip le fichier docx pour avoir le xml
        document = zipfile.ZipFile(file_docx)     

        # recupere le xml avec l'encodage utf-8
        content = document.read('word/document.xml').decode('utf-8')

        #print(content)

        # 'xml' is the parser used. For html files, which BeautifulSoup is typically used for, it would be 'html.parser'.
        soup = BeautifulSoup(content, 'xml')

        return soup

    except FileNotFoundError:

        print("file not found.") 


def xml_to_hyperlink(soup:BeautifulSoup, file_name) :  

    # balise XML pour lien <w:hyperlink></w:hyperlink>
    hyperlink = soup.find_all('w:hyperlink')

    hyperlink_text = {}
    hyperlink_rels = {}


    for name in hyperlink:
        hyperlink_text[name.get('r:id')] = name.text    

    rels = Document(file_name+".docx").part.rels

    for rel in rels:
        if rels[rel].reltype == RT.HYPERLINK:  
            hyperlink_rels[rel] = rels[rel]._target

    print(hyperlink_rels)

    hyperkink_merge = {}
    h_text = []
    h_link = []

    # fusionner text et rels
    for cle, valeur in hyperlink_rels.items():
        print("l'élément de clé", cle, "vaut", valeur)
        hyperkink_merge[cle] = {"text":hyperlink_text[cle],"link":valeur}
        h_text.append(hyperlink_text[cle])
        h_link.append(valeur)
        

    #HYPERLINK_MERGE.append(hyperkink_merge)
    with_file_name_as_root = {file_name:hyperkink_merge}
    HYPERLINK_MERGE.append(with_file_name_as_root)

    HYPERLINK_TEXT.append(h_text)
    HYPERLINK_LINK.append(h_link)


def hyperlink_to_json(hyperlink) :
    return json.dumps(hyperlink)

def save_json_file(file_json, hyperlink) :

    f = open(file_json,"w")
    f.truncate(0)
    f.write(hyperlink_to_json(hyperlink))
    f.close()


def draw_graph() :    

    FILES_NAMES = []
    FILES_NAMES_JSON = []
    FILES_NAMES_DOCX = []

    # On parcours tout les fichiers à traiter

    i = 0

    for f in FILES_NAMES :
        #file_name = input("Le nom du fichier sans l'extension .docx :")  
        if(file_to_xml(f) != None) :
            xml = file_to_xml(f)
            xml_to_hyperlink(xml,f)
            f_json = hyperlink_to_json(HYPERLINK_MERGE[i])
            save_json_file(f+".json", HYPERLINK_MERGE[i])
        else :
            print("erreur : vérifier le nom du fichier et son enplacement.")
        i+=1        


    #print(FILES_NAMES)
    #print(FILES_NAMES_DOCX)
    #print(FILES_NAMES_JSON)
    print(HYPERLINK_MERGE)
    print(HYPERLINK_LINK)
    print(HYPERLINK_TEXT)

    # edges_labels = {}    

    i = 0
    # on parcours tous les documents pour construire le graphe
    for f in FILES_NAMES_DOCX :
        j = 0
        # parcours tous les liens du document courant
        for link in HYPERLINK_LINK[i] :        
            # si la chaine se termine par .docx (renvoie sur un autre docx)
            if HYPERLINK_LINK[i][j].endswith(".docx") :
                FILES_RELATIONS.append((f,HYPERLINK_LINK[i][j]))             
                # edges_labels[(f,HYPERLINK_LINK[i][j])] = HYPERLINK_TEXT[i][j]                           
            # liens externes            
            elif not HYPERLINK_LINK[i][j].endswith(".docx"):
                FILES_RELATIONS.append((f,"LIEN EXTERNE"))            
                # edges_labels[(f,"LIEN EXTERNE")] = (HYPERLINK_TEXT[i][j],HYPERLINK_LINK[i][j])         
            # print(other_links)        
            j+=1
        i+=1


    # determiner le poids des relations
    i = 0
    for r in FILES_RELATIONS :
        FILES_WEIGHT[FILES_RELATIONS[i]] = FILES_RELATIONS.count(FILES_RELATIONS[i])
        i+=1


    print(FILES_WEIGHT)


    G = nx.DiGraph()
    G.add_edges_from(FILES_RELATIONS)

    pos = nx.spring_layout(G)
    plt.figure()
    nx.draw(
        G, pos, edge_color='black', width=1, linewidths=1,
        node_size=500, node_color='cyan', alpha=0.9,
        labels={node: node for node in G.nodes()}
    )
    nx.draw_networkx_edge_labels(
        G, pos,
        edge_labels=FILES_WEIGHT,
        font_color='blue'
    )
    plt.axis('off')
    plt.show()


# sauvegarde les documents et leurs liens dans une base de donnée MySQL
def save_to_db() :

    i = 0

    for f in FILES_NAMES :
        if(file_to_xml(f) != None) :
            xml = file_to_xml(f)
            xml_to_hyperlink(xml,f)
            f_json = hyperlink_to_json(HYPERLINK_MERGE[i])
            save_json_file(f+".json", HYPERLINK_MERGE[i])
        else :
            print("erreur : vérifier le nom du fichier et son enplacement.")
        i+=1            

    try:
        connection = mysql.connector.connect(host=DATABASE_HOST,
                                         database=DATABASE,
                                         user=DATABASE_USER,
                                         password=DATABASE_PASSWORD)

        
        if connection.is_connected():
            db_Info = connection.get_server_info()
            print("Connected to MySQL Server version ", db_Info)

            
            # cursor.execute("select database();")

            insert_file = "INSERT INTO fichier (nom_f) VALUES (%s)"            
            insert_link = ("INSERT INTO lien" "(fichier_source, fichier_cible, direction, contenu)" "VALUES (%s, %s, %s, %s)")

            select_fichier_source = "SELECT max(id_f) FROM fichier"

            fichier_source_id = []

            # ajout à la table FICHIER
            for n in FILES_NAMES_DOCX :                
                cursor = connection.cursor()
                # Since you are using mysql module, cursor.execute requires a sql query and a tuple as parameters
                cursor.execute(insert_file, (n,))
                print("fichier ajouté à la table fichier :", n)
                connection.commit()                

                # sauvegarde l'id des fichiers ajoutés à la base
                cursor.execute(select_fichier_source)
                res = cursor.fetchall()
                for r in res :
                    print("r value",r)
                    fichier_source_id.append(int(''.join(filter(str.isdigit,str(r)))))

                            
            #print("fichier_source_id",fichier_source_id)
            
            # fichier_cible_possible =  ''.join([chr(x) for x in fichier_source_id])        
            fichier_cible_possible = str(fichier_source_id)[1:-1]
            fichier_cible_possible = tuple(map(int, fichier_cible_possible.split(', ')))      

            #print("fichier_cible_possible",fichier_cible_possible)


        i = 0
        # on parcours tous les documents pour construire le graphe
        for f in FILES_NAMES_DOCX :            
            j = 0
            # parcours tous les liens du document courant
            for link in HYPERLINK_LINK[i] :                          
                cursor = connection.cursor()
                # si la chaine se termine par .docx (renvoie sur un autre docx)
                if HYPERLINK_LINK[i][j].endswith(".docx") :                         
                    # determiner l'id du fichier cible avec son nom
                    FILES_RELATIONS.append((f,HYPERLINK_LINK[i][j]))
                   
                    # en parametres le nom du fichier cible & un id parmis ceux ajouté (très important si fichiers de même nom)               
                    a = "SELECT id_f FROM fichier WHERE nom_f="+'"'+HYPERLINK_LINK[i][j]+'"'+" AND id_f IN "+str(fichier_cible_possible)

                    cursor.execute(a)               
                    res = cursor.fetchall()

                    for r in res :                        
                        cursor.execute(insert_link, (fichier_source_id[i],str(r[0]),HYPERLINK_LINK[i][j],HYPERLINK_TEXT[i][j]))
                        print("table lien MAJ")
                    connection.commit()
                # liens externes            
                elif not HYPERLINK_LINK[i][j].endswith(".docx"):
                    FILES_RELATIONS.append((f,"LIEN EXTERNE"))                    
                    cursor.execute(insert_link, (fichier_source_id[i],None,HYPERLINK_LINK[i][j],HYPERLINK_TEXT[i][j]))
                    print("table lien MAJ")
                    connection.commit()
                # print(other_links)        
                j+=1
            i+=1


    except Error as e:
        print("Error while connecting to MySQL", e)
    finally:
        if connection.is_connected():
            cursor.close()
            connection.close()
            print("MySQL connection is closed") 



# tracer un graphe avec les informations de la bdd 
# params : id_fichiers (tableau des fichiers)
# 1 - se connecter à la db
# 2 - extraire le(s) fichier(s) sélectionné(s)
# 3 - tracer le graphe
def draw_graph_from_db(id_fichiers) :

    try:
        connection = mysql.connector.connect(host=DATABASE_HOST,
                                         database=DATABASE,
                                         user=DATABASE_USER,
                                         password=DATABASE_PASSWORD)

        if connection.is_connected():
            db_Info = connection.get_server_info()
            print("Connected to MySQL Server version ", db_Info)

            # les requêtes utilisées
            select_filename = "SELECT nom_f from FICHIER WHERE id_f=%s"
            select_all_link = "SELECT * from LIEN WHERE fichier_source=%s"

            FILES_NAMES_DOCX = []        
            
            # tableau double dimension avec tous les liens
            all_link_sort = []         

            # parcourrir tout les id des fichiers sources
            for id in id_fichiers :
                print("B", id)
                cursor = connection.cursor(buffered=True)
                # Since you are using mysql module, cursor.execute requires a sql query and a tuple as parameters
                cursor.execute(select_filename, (id,))
                res = cursor.fetchall()
                for r in res :
                    FILES_NAMES_DOCX.append(r[0])  

                connection.commit() 
                cursor.execute(select_all_link, (id,))
                res = cursor.fetchall()
                
                all_link_sort.append(res)

            print("DDDDDDD",all_link_sort)                                
            print("A",FILES_NAMES_DOCX)


            # determiner les relations du graphe
            i = 0
            for fn in FILES_NAMES_DOCX :
                j = 0
                for l in all_link_sort[i] :
                    # si la cible se termine par .docx (renvoie sur un autre docx)
                    if l[3].endswith(".docx") :
                        FILES_RELATIONS.append(("#"+str(id_fichiers[i]+" | "+str(fn)),"#"+str(l[2])+" | "+str(l[3])))             
                        # edges_labels[(f,HYPERLINK_LINK[i][j])] = HYPERLINK_TEXT[i][j]                           
                    # liens externes            
                    elif not l[3].endswith(".docx"):
                        FILES_RELATIONS.append(("#"+str(id_fichiers[i]+" | "+str(fn)),"LIEN EXTERNE"))            
                        # edges_labels[(f,"LIEN EXTERNE")] = (HYPERLINK_TEXT[i][j],HYPERLINK_LINK[i][j])         
                        # print(other_links)    
                    j+=1
                i+=1


        # determiner le poids des relations
        i = 0
        for r in FILES_RELATIONS :
            FILES_WEIGHT[FILES_RELATIONS[i]] = FILES_RELATIONS.count(FILES_RELATIONS[i])
            i+=1

        G = nx.DiGraph()
        G.add_edges_from(FILES_RELATIONS)

        pos = nx.spring_layout(G)
        plt.figure()
        nx.draw(
            G, pos, edge_color='black', width=1, linewidths=1,
            node_size=500, node_color='cyan', alpha=0.9,
            labels={node: node for node in G.nodes()}
        )
        nx.draw_networkx_edge_labels(
            G, pos,
            edge_labels=FILES_WEIGHT,
            font_color='blue'
        )
        plt.axis('off')
        plt.show()            


    except Error as e:
        print("Error while connecting to MySQL", e)
    finally:
        if connection.is_connected():
            # cursor.close()
            connection.close()
            print("MySQL connection is closed") 

    return 0
















arret = False

while not (arret) :

    print("[INFO] Les fichiers actuels :"+str(FILES_NAMES))
    print("[0] Ajouter un fichier")
    print("[1] Ne plus ajouter, tracer le graphe")
    print("[2] Tracer le graphe à partir de la base de donnée")

    try : 

        choix = input()

        if (int(choix)==0) :    
            file_name = input("Le nom du fichier sans l'extension .docx :")  
            if(os.path.exists(file_name+".docx")) :
                FILES_NAMES.append(file_name)
                file_docx =file_name+".docx"
                file_json = file_name+".json"
                FILES_NAMES_DOCX.append(file_docx)
                FILES_NAMES_JSON.append(file_json)
            else :
                print("[ERREUR] Le fichier n'existe pas\n")    
        elif (int(choix)==1) :
            print("[INFO] Les fichiers traités :"+str(FILES_NAMES))
            save_to_db()
            draw_graph()            
            arret = True 
        elif (int(choix)==2) :
            file_id = input("Id des fichiers traités sous la forme id1,id2,id3,...,idn :")
            print("A",file_id)
            arr = str(file_id).split(',')
            draw_graph_from_db(arr)            
            arret = True    
    except ValueError :
        print("[INFO] Choix invalide")        


 