cell-label
==========

Website for labeling XL cells


## Updating the XLS files

1. Add xls files to corpus skyDrive folder
2. Make them all public (otherwise, the file token will require an authid which we cannot obtain programnmatically)
3. In http://isdk.dev.live.com/dev/isdk/ISDK.aspx?category=scenarioGroup_core_concepts&index=0 run code:

```
WL.init({ client_id: clientId, redirect_uri: redirectUri });

WL.login({ "scope": "wl.skydrive" }).then(
    function(response) {
        getFiles();
    },
    function(response) {
        log("Could not connect, status = " + response.status);
    }
);

function getFiles() {
    var files_path = "folder.072e74b1abfc5464.72E74B1ABFC5464!190/files";
    WL.api({ path: files_path, method: "GET" }).then(
        onGetFilesComplete,
        function(response) {
            log("Cannot get files and folders: " +
                JSON.stringify(response.error).replace(/,/g, ",\n"));
        }
    );
}

function onGetFilesComplete(response) {
    var items = response.data;
    for (var i = 0; i < items.length; i++) {
        log(items[i].name + "#" + items[i].id)
               // JSON.stringify(items[i]).replace(/,/g, ",\n"));
       // if (items[i].type === "folder") {
       //     log("Found a folder with the following information: " +
       //         JSON.stringify(items[i]).replace(/,/g, ",\n"));
       //     foundFolder = 1;
       //     break;
       // }
    }
}

function log(message) {
    var child = document.createTextNode(message);
    var parent = document.getElementById('JsOutputDiv') || document.body;
    parent.appendChild(child);
    parent.appendChild(document.createElement("br"));
}
```

4. Place result in input.txt