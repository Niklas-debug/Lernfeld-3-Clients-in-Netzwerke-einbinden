from flask import Flask, redirect, url_for, render_template, request
from pyad import *
import pythoncom
from flask_basicauth import BasicAuth

app = Flask(__name__)

app.config['BASIC_AUTH_USERNAME'] = 'Administrator'
app.config['BASIC_AUTH_PASSWORD'] = 'BS14Lf3'
basic_auth = BasicAuth(app)
app.config['BASIC_AUTH_FORCE'] = True
app.config['BASIC_AUTH_REALM'] = "Bitte den Benutzernamen und das Passwort vom GEEK-DC-01 eingeben."

@app.route("/")
def home():
	return render_template("index.html")

@app.route("/features")
def featuresSite():
    return render_template("features.html")

@app.route("/team")
def teamSite():
    return render_template("team.html")


def replaceUmlauts(string):
    string = string.replace('ü', 'u')
    string = string.replace('ö', 'o')
    string = string.replace('ä', 'a')
    string = string.replace('Ü', 'U')
    string = string.replace('Ö', 'O')
    string = string.replace('Ä', 'A')
    return string


def createUser(givenName, sn, department, pas):
    pythoncom.CoInitialize()
    ou = pyad.adcontainer.ADContainer.from_dn("OU=Ben_" + department + ",OU=Benutzer,OU=GEEK-Fitness,DC=geek,DC=local")
    sAMAccountName = replaceUmlauts(sn)[0:4].lower() + replaceUmlauts(givenName)[0:2].lower()
    displayName = givenName + " " + sn
    userPrincipalName = sAMAccountName + "@geek.local"
    profilePath = r"\\geek-dc-1\Profile$\%username%"
    homeDrive = "H:"
    homeDirectory = r"\\geek-dc-1\Homes$\%username%"
    distinguishedName = "CN="+displayName+",OU=Ben_"+department+",OU=Benutzer,OU=GEEK-Fitness,DC=geek,DC=local"
    new_user = pyad.aduser.ADUser
    new_user = pyad.aduser.ADUser.create(displayName, ou, password=pas, upn_suffix=None, enable=True, optional_attributes={"userPrincipalName": userPrincipalName,"sAMAccountName": sAMAccountName,"givenName": givenName,"displayName": displayName,"sn": sn,"homeDirectory": homeDirectory,"homeDrive": homeDrive})
    group = pyad.adgroup.ADGroup.from_cn("Grp_"+department)
    rdpGroup = pyad.adgroup.ADGroup.from_cn("Remotedesktopbenutzer")
    group.add_members([new_user])
    rdpGroup.add_members([new_user])
    return True


@app.route("/addUser", methods=["POST", "GET"])
def addUser():
	if request.method == "POST":
		vorname = request.form["vorname"]
		nachname = request.form["nachname"]
		abteilung = request.form["category"]
		pas = request.form["passwort"]

		if len(vorname) == 0:
			return redirect(url_for("addUser", error="Der Vorname darf nicht leer sein!"))
			exit(1)
		elif any(chr.isdigit() for chr in vorname):
			return redirect(url_for("addUser", error="Der Vorname darf keine Zahlen enthalten!"))
			exit(1)
		elif len(nachname) == 0:
			return redirect(url_for("addUser", error="Der Nachname darf nicht leer sein!"))
			exit(1)
		elif any(chr.isdigit() for chr in nachname):
			return redirect(url_for("addUser", error="Der Nachname darf keine Zahlen enthalten!"))
			exit(1)
		elif len(pas) < 6:
			return redirect(url_for("addUser", error="Das Passwort darf nicht weniger als 6 Zeichen haben!"))
			exit(1)

		if createUser(vorname,nachname,abteilung,pas):
		    return redirect(url_for("addUser", success="Dieser Benutzer wurde erfolgreich erstellt und besitzt ggf. die Richtige Gruppe/Berechtigungen."))
		else:
		    return redirect(url_for("addUser", error="Dieser Benutzer wurde nicht erfolgreich erstellt."))
	else:
		return render_template("addUser.html")


@app.route("/deleteUser", methods=["POST", "GET"])
def delUser():
    if request.method == "POST":
        employeeName = request.form["employeeName"]
        try:
            pythoncom.CoInitialize()
            pyad.aduser.ADUser.from_cn(employeeName).delete()
            return redirect(url_for("delUser", success="Dieser Benutzer wurde erfolgreich gelöscht."))
            return True
        except:
            return redirect(url_for("delUser", error="Dieser Benutzer konnte nicht erfolgreich gelöscht werden."))
            return False
    else:
        return render_template("deleteUser.html")


@app.route("/changeDepartment", methods=["POST", "GET"])
def changeDep():
    if request.method == "POST":
        try:
            pythoncom.CoInitialize()
            employeeName = request.form["employeeName"]
            departmentOld = request.form["category1"]
            departmentNew = request.form["category2"]

            userGroup = aduser.ADUser.from_cn(employeeName)
            newDepartment = adgroup.ADGroup.from_cn("Grp_"+departmentOld)
            newDepartment.remove_members(userGroup)

            userOU = aduser.ADUser.from_cn(employeeName)
            userOU.move(adcontainer.ADContainer.from_dn("OU=Ben_"+departmentNew+",OU=Benutzer,OU=GEEK-Fitness,DC=geek,DC=local"))

            userGroup = aduser.ADUser.from_cn(employeeName)
            newDepartment = adgroup.ADGroup.from_cn("Grp_"+departmentNew)
            newDepartment.add_members(userGroup)

            return redirect(url_for("changeDep", success="Die Abteilung vom Benutzer wurde erfolgreich geändert."))
            return True
        except:
            return redirect(url_for("changeDep", error="Die Abteilung vom Benutzer konnte nicht erfolgreich geändert werden."))
            return False
    else:
        return render_template("changeDepartment.html")


@app.route("/changePermission", methods=["POST", "GET"])
def changePer():
    if request.method == "POST":
        try:
            pythoncom.CoInitialize()
            employeeName = request.form["employeeName"]
            addOrRemove = request.form["category1"]
            department = request.form["category2"]

            if addOrRemove == "addPerp":
                userGroup = aduser.ADUser.from_cn(employeeName)
                newDepartment = adgroup.ADGroup.from_cn("Grp_"+department)
                newDepartment.add_members(userGroup)
            elif addOrRemove == "rmPerp":
                userGroup = aduser.ADUser.from_cn(employeeName)
                newDepartment = adgroup.ADGroup.from_cn("Grp_"+department)
                newDepartment.remove_members(userGroup)

            return redirect(url_for("changePer", success="Die Berechtigung vom Benutzer, für den Ordner wurde erfolgreich geändert."))
            return True
        except:
            return redirect(url_for("changePer", error="Die Berechtigung vom Benutzer, für den Ordner konnte nicht erfolgreich geändert werden."))
            return False
    else:
        return render_template("changePermission.html")


if __name__ == "__main__":
	app.run(debug=True)
