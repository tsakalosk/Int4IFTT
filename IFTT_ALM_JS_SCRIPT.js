// Int4 IFTT integration
//
//
//
//
//
//
//
//
//
//
// ====================================================
function getPasswordViaPrompt() {
    var oIE = new ActiveXObject("InternetExplorer.Application");
    oIE.Navigate("about:blank");

    oIE.Document.Title = "Form to insert Password";
    oIE.ToolBar = false;
    oIE.Resizable = false;
    oIE.StatusBar = false;
    oIE.Width = 320;
    oIE.Height = 180;

    //Insert the HTML code to prompt for a password
    oIE.document.body.innerHTML =
        "<table>" +
        "<tr>" +
        " <td>User name</td>" +
        "<td><input type='text' id='User' value='" + TDConnection.UserName + "'></td>" +
        "</tr>" +
        "<tr>" +
        "  <td>Password</td>" +
        "<td><input type='password' size='20' id='Password'/></td>" +
        " </tr>" +
        " </table>" +
        "<input type='submit' value='OK' onclick='document.getElementById(\"OK\").value = 1'>" +
        "<input type='hidden' id='OK' name='OK' value='0'>";

    oIE.Visible = true;

    try {
        do {
            XTools.Sleep(100);

            var ok = oIE.document.getElementById("OK").value;

        } while (ok !== '1');
    } catch (err) {}

    var oReturn = {
        user: oIE.document.getElementById("User").value,
        password: oIE.document.getElementById("Password").value
    }

    oIE.Visible = false;
    oIE.Quit();
    oIE = null;

    return oReturn;
}

function attachResponse(responseText, currentRun) {

    var fso = new ActiveXObject("Scripting.FileSystemObject");

    //get Temp folder
    var path = fso.GetSpecialFolder(2);
    var timeInMs = new Date().getTime();
    path += "\\" + "iftt_response_" + timeInMs + ".xml";
    //TDOutput.Print(path);
    var file = fso.CreateTextFile(path, true);
    file.Write(responseText);
    file.Close();
    var objRunAttach = currentRun.Attachments;
    var newAttach = objRunAttach.AddItem(null);
    newAttach.FileName = path;

    newAttach.Type = 1;
    newAttach.Post();
    fso.DeleteFile(path);
}

function attachUrl(url, currentRun) {
    var objRunAttach = currentRun.Attachments;
    var newAttach = objRunAttach.AddItem(null);
    newAttach.FileName = url;

    newAttach.Type = 2;
    newAttach.Post();
}

function getRequest(testRunName, parameterValueFactory) {

    var list = parameterValueFactory.NewList("");
    var par = list.Item(1);
    TDOutput.Print(par.Name + ":" + par.ActualValue);
    var testCases = par.ActualValue.split(",");
    var itTestCasesListItemsStr = "";
    for (var i = 0; i < testCases.length; i++) {
        if (testCases[i] !== "") {
            itTestCasesListItemsStr += '<item>' + testCases[i] + '</item>';
        }
    }

    if (itTestCasesListItemsStr !== "") {
        itTestCasesListItemsStr = '<ItTestCasesList>' + itTestCasesListItemsStr + '</ItTestCasesList>';
    }

    var par = list.Item(2);
    TDOutput.Print(par.Name + ":" + par.ActualValue);
    var testScenarios = par.ActualValue.split(",");
    var ItScenariosListItemsStr = "";
    for (i = 0; i < testScenarios.length; i++) {
        if (testScenarios[i] !== "") {
            ItScenariosListItemsStr += '<item>' + testScenarios[i] + '</item>';
        }
    }

    if (ItScenariosListItemsStr !== "") {
        ItScenariosListItemsStr = '<ItScenariosList>' + ItScenariosListItemsStr + '</ItScenariosList>';
    }

    var request = '<?xml version="1.0" encoding="utf-8"?>' +
        '<soapenv:Envelope xmlns:soapenv="http://schemas.xmlsoap.org/soap/envelope/" xmlns:urn="urn:sap-com:document:sap:soap:functions:mc-style">' +
        '<soapenv:Header/>  ' +
        '<soapenv:Body>  ' +
        '<urn:RunTests>   ' +
        ItScenariosListItemsStr +
        itTestCasesListItemsStr +
        '<IvTestRunName>' + testRunName + '</IvTestRunName>  ' +
        '</urn:RunTests>   ' +
        '</soapenv:Body>  ' +
        '</soapenv:Envelope>';

    return request;
}

// ----------------------------------------------------
// Main Test Function
// Debug - Boolean. Equals to false if running in [Test Mode] : reporting to Application Lifecycle Management
// CurrentTestSet - [OTA COM Library].TestSet.
// CurrentTSTest - [OTA COM Library].TSTest.
// CurrentRun - [OTA COM Library].Run.
// ----------------------------------------------------
function Test_Main(Debug, CurrentTestSet, CurrentTSTest, CurrentRun) {
    try {
        TDConnection.IgnoreHtmlFormat = true;
        TDOutput.Clear();
        TDOutput.Print("Start IFTT Test");
        var testRunName = "ALM Run ID: " + CurrentRun.ID; ;
        TDOutput.Print(testRunName);
        var IFTT_CONFIG = "int4 IFTT";

        // Option 1 - user , password popup
        var credentials = getPasswordViaPrompt();
        var username = credentials.user;
        var password = credentials.password;
        var ifttUrl = "http://server/sap/bc/srt/scs/int4/iftt_run_tests?sap-client=001";

        // Option 2 - user, password from Project Lists
        //var username = TDConnection.Customization.Lists.List(IFTT_CONFIG).RootNode.Child("wsUser").Children.Item(1).Name;
        //var password = TDConnection.Customization.Lists.List(IFTT_CONFIG).RootNode.Child("wsPassword").Children.Item(1).Name;
        //var ifttUrl = TDConnection.Customization.Lists.List(IFTT_CONFIG).RootNode.Child("wsUrl").Children.Item(1).Name;

        //Option 3 ( for testing) - user, password in script
        //var username = "user";
        //var password = "password" ;
        //var ifttUrl = "https://server/sap/bc/srt/scs/int4/iftt_run_tests?sap-client=001";


        var xhr = new ActiveXObject("Microsoft.XMLHTTP");
        var body = getRequest(testRunName, CurrentTSTest.ParameterValueFactory);

        xhr.open("POST", ifttUrl, false, username, password);
        xhr.setRequestHeader("SOAPAction", "urn:sap-com:document:sap:soap:functions:mc-style:_-INT4_-IFTT_RUN_TESTS:RunTestsRequest");
        xhr.setRequestHeader("Content-Type", "text/xml;charset=UTF-8");

        TDOutput.Print("Sending request to " + ifttUrl + " ...");
        xhr.send(body);
        TDOutput.Print("Checking response...");
        if (!Debug && xhr.status !== 200) {
            CurrentRun.Status = "Failed";
            CurrentTSTest.Status = "Failed";
        }

        TDOutput.Print("HTTP Status:" + xhr.status);

        var error = xhr.responseText.match(/<EvErrorMessage>(.+?)<\/EvErrorMessage>/);
        if (error !== null) {
            TDOutput.Print(error[1]);
        }

        var url = xhr.responseText.match(/<EvReportUrl>(.+?)<\/EvReportUrl>/);
        if (url !== null) {
            url = url[1].replace(/&amp;/g, '&');
            attachUrl(url, CurrentRun);
        }
        var passed = xhr.responseText.match(/<EvTcPassed>(.+?)<\/EvTcPassed>/)[1]; ;
        var failed = xhr.responseText.match(/<EvTcFailed>(.+?)<\/EvTcFailed>/)[1]; ;
        var total = xhr.responseText.match(/<EvTcTotal>(.+?)<\/EvTcTotal>/)[1]; ;

        TDOutput.Print("Total IFTT test cases: " + total);
        TDOutput.Print("Passed IFTT test cases: " + passed);
        TDOutput.Print("Failed IFTT test cases: " + failed);
        TDOutput.Print("Report url: " + url);

        attachResponse(xhr.responseText, CurrentRun);


        if (!Debug) {
            var step = CurrentRun.StepFactory.addItem(null);
            step.Name = "IFTT";

            if (passed > 0 && failed == 0) {
                TDOutput.Print("Passed");
                step.Status = "Passed";
                //  CurrentRun.Status = "Passed";
                // CurrentTSTest.Status = "Passed";
            } else {
                step.Status = "Failed";
            }

            step.Field("ST_ACTUAL") = TDOutput.Text;
            step.Post();
        }
    }
    // handle run-time errors
    catch (e) {
        TDOutput.Print("Run-time error [" + (e.number & 0xFFFF) + "] : " + e.description);
        TDOutput.Print(e.stack);
        TDOutput.Print(e.name);
        TDOutput.Print(e.message);
        // update execution status in "Test" mode
        if (!Debug) {
            CurrentRun.Status = "Failed";
            CurrentTSTest.Status = "Failed";
        }
    }
}