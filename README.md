# Eines OAC — Web (prova de concepte)

Versió **web** del Corregidor d'Àmbits que s'executa **íntegrament al navegador**
amb [Pyodide](https://pyodide.org) (Python compilat a WebAssembly).

## Per què aquesta arquitectura

Les dades policials **mai no surten de l'ordinador de l'usuari**: l'Excel
seleccionat es processa dins del navegador i el resultat es descarrega localment.
El servidor (o GitHub Pages) només serveix *codi*, mai *dades*. Això evita el
problema de protecció de dades que tindria pujar fitxers a un servidor extern.

## Fitxers

| Fitxer | Funció |
|--------|--------|
| `index.html` | Interfície web + orquestració de Pyodide |
| `corregidor_core.py` | Lògica de correcció (extreta de `corregidor.py`, sense Tkinter) |
| `logo_mossos.png` | Logo per a la capçalera i l'Excel de sortida |
| `servir_local.bat` | Aixeca un servidor local per provar |

## Com provar-ho en local

Pyodide **necessita HTTP** (no funciona obrint `index.html` amb doble clic).

1. Doble clic a `servir_local.bat` (o, a la carpeta, `python -m http.server 8000`).
2. Obre el navegador a **http://localhost:8000**.
3. Espera que carregui el motor (uns segons la primera vegada).
4. Arrossega un Excel d'àmbits i prem **Processar**.
5. Descarrega l'Excel corregit i compara'l amb el de la versió d'escriptori.

## Prova des d'un ordinador corporatiu

L'objectiu és confirmar que una màquina corporativa pot:
- Carregar Pyodide (des del CDN `jsdelivr` o autoallotjat).
- Executar WebAssembly al navegador.

Si el CDN està bloquejat, caldrà **autoallotjar Pyodide** (descarregar els fitxers
i servir-los des de la mateixa web). Ho veurem segons el resultat de la prova.

## Desplegament públic (futur)

En estar tot al client, es pot allotjar de franc a **GitHub Pages** sense que cap
dada hi passi mai. Pendent de validar a l'entorn corporatiu.
