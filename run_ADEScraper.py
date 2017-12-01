import sys, os

path = ""
out = ""
launchCommand = ""

if ("--arboscrap" in sys.argv or "-as" in sys.argv):
    # lancer le script 'arboScraper'

    launchCommand += "ArboScraper.exe"

    if (len(sys.argv) > 3):
        path = sys.argv.pop(sys.argv.index("--arboscrap" if "--arboscrap" in sys.argv else "-as") + 1)
        if (path.startsWith("-")):
            raise Exception("ERROR - Invalid path : ", path)

elif (len(sys.argv) >= 3 and ("--pathscrap" in sys.argv or "-ps" in sys.argv)):
    # lancer le script 'pathScraper' avec comme path l'arg suivant

    launchCommand += "PathScraper.exe"

    path = sys.argv.pop(sys.argv.index("--pathscrap" if "--pathscrap" in sys.argv else "-ps") + 1)

    if (path.startsWith("-")):
        raise Exception("ERROR - Invalid path : " + str(path))


launchCommand += " " + path


for i, arg in enumerate(sys.argv[1:]):
    if (arg == "--help" or arg == "-h" or arg == "?"):
        print("HELP :",
              "    --arboscrap (ou -as) [path] : scrap l'arborescence à partir de 'path' ou entièrement si 'path' n'est pas précisé\n",
              "    --pathscrap (ou -ps) path   : scrap les fichiers suivant le 'path' précisé\n",
              "    --out       (out -o) file   : définit 'file' comme étant le fichier où les données seront enregistrées"
              "    --debug     (ou -d)         : lance le script en admettant que ADE et le débuggeur sont déjà ouverts\n",
              "    --startup   (ou -s)         : force le script à ouvrir ADE au démarrage\n",
              "    --append    (ou -a)         : ajouter des lignes au fichier de sortie au lieu de l'effacer\n",
              "tous les 'path' doivent faire référence à un fichier où se trouve le script, de la forme 'nomDuPath.txt'\n")

    elif (arg == "--debug" or "-d"):
        launchCommand += " --debug"

    elif (arg == "--startup" or "-s"):
        launchCommand += " --startup"

    elif (arg == "--append" or "-a"):
        launchCommand += " --append"

    elif (arg == "--out" or "-o"):
        if i + 1 <= len(sys.argv) -1:
            out = sys.argv[i]
            if out.startswith("-"):
                raise Exception("ERROR - Invalid output file :" + str(out))
        else:
            raise Exception("ERROR - No output file specified.")


if not(launchCommand.startswith(" ")):
    print("Starting the script...")
    print("command sent :", launchCommand)
    os.system(launchCommand)

elif len(sys.argv) > 2:
    raise Exception("ERROR - Incorrect parameters : '" + str(launchCommand) + "'")