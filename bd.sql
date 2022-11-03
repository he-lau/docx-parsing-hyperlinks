CREATE TABLE FICHIER (
    id_f INT NOT NULL AUTO_INCREMENT,
    nom_f VARCHAR(255) NOT NULL,

    PRIMARY KEY (id_f)    
);

CREATE TABLE LIEN (
    id_l INT NOT NULL AUTO_INCREMENT,
    fichier_source INT NOT NULL REFERENCES FICHIER(id_f),
    fichier_cible INT REFERENCES FICHIER(id_f),
    direction VARCHAR(255) NOT NULL,
    contenu VARCHAR(255) NOT NULL,

    PRIMARY KEY (id_l)    
);