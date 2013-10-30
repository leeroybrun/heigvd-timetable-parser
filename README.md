# HEIG-VD timetable parser

Ce script python sert à transformer un horaire au format XLS provenant de l'intranet du département FEE de la HEIG-VD en un fichier ICS.
Ce fichier peut ensuite être facilement importé dans n'importe quel logiciel de gestion de calendrier (Outlook, Gmail, etc).

## Installation

Il vous suffit de lancer le script `setup.py` afin d'installer les dépendances du projet.

```shell
python setup.py install
```

Utilisez [virtualenv](https://pypi.python.org/pypi/virtualenv) si vous ne souhaitez pas que les dépendances soit installées globalement sur votre système.

## Utilisation

Le script s'utilise comme suit :

```shell
python main.py fichierSource [filtre]
```

Il vous suffit donc de récupérer le fichier d'horaires sur l'Intranet de la HEIG-VD puis de le placer dans le même dossier que ce script.

Vous pouvez ensuite lancer le script comme ceci (si votre fichier s'appelle `forIL_1.xls`) :

```shell
python main.py forIL_1.xls
```

Un fichier `forIL_1.ics` sera alors créé dans le même répertoire que le fichier source.

## Filtres

Les filtres vous permettent de n'exporter qu'un certain type d'événement. Par défault, toutes les informations du fichier Excel sont exportées dans un seul et même fichier (cours, tests, examens, congés).

Les filtres disponibles sont les suivants :

- `course` : Exporte uniquement les jours de cours
- `test` : Exporte uniquement les tests
- `exam` : Exporte uniquement les examens
- `test_exam` : Exporte uniquement les tests et examens supperposés
- `free` : Exporte uniquement les congés

Pour n'exporter que les examens, il suffit donc de lancer le script comme ceci :

```shell
python main.py forIL_1.xls exam
```

Un fichier `forIL_1_exam.ics` sera alors créé dans le même répertoire que le fichier source.

Licence
======================
(The MIT License)

Copyright (C) 2013 Leeroy Brun, www.leeroy.me

Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the "Software"), to deal in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.

[![Bitdeli Badge](https://d2weczhvl823v0.cloudfront.net/leeroybrun/heigvd-timetable-parser/trend.png)](https://bitdeli.com/free "Bitdeli Badge")
