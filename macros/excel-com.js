
var COMMANDS = {
    csvout: cmd_csvout
};

var subcmd = WScript.Arguments(0);
if (COMMANDS[subcmd]) {
    var Excel = new ActiveXObject('Excel.Application');
    try {
        COMMANDS[subcmd].apply(null, convertArgs(WScript.Arguments));
    } finally {
        // Clean-up
        var dontsave = 0;
        for (var i = 1; i <= Excel.Workbooks.Count; i++) {
            Excel.Workbooks(i).Close(dontsave);
        }
        Excel.Quit();
    }
    WScript.Sleep(3000);
}
else {
    WScript.Echo("Error: Unknown sub-command '" + subcmd + "'.");
    WScript.Sleep(3000);
    WScript.Quit(1);
}



function cmd_csvout(src, dest) {
    var oFSO = new ActiveXObject('Scripting.FileSystemObject');
    try {
        oFSO.DeleteFile(dest);
    } catch (e) {}

    var xlCSV = 6;
    Excel.Workbooks.Open(src).SaveAs(dest, xlCSV);
}

function convertArgs(wsargs) {
    var args = [];
    // Dispose first subcommand argument.
    for (var i = 1; i < wsargs.length; i++) {
        args.push(wsargs(i));
    }
    return args;
}
