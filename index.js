// Starts the connection
console.log("Test2");
async function run() {
  console.log("Connect...");
  const msalConfig = {
    auth: {
      clientId: "b095fc0c-93a9-441a-aa2f-e34a27942e5d",
      authority:
        "https://login.microsoftonline.com/d4616c26-b9bd-4d02-91d7-60ea7be3789a/",
      redirectUri: "http://localhost:8080",
    },
  };
  const msalClient = new msal.PublicClientApplication(msalConfig);

  const authProvider =
    new MSGraphAuthCodeMSALBrowserAuthProvider.AuthCodeMSALBrowserAuthenticationProvider(
      msalClient,
      {
        account: {}, // the AccountInfo instance to acquire the token for
        scopes: ["user.read", "Directory.ReadWrite.All"],
        interactionType: msal.InteractionType.Popup,
      }
    );

  const graphClient = MicrosoftGraph.Client.initWithMiddleware({
    authProvider,
  });

  // outputs /me Get api request
  const profile = await graphClient.api("/me").get();
  console.dir(profile);

  const ImportGroups = [
    "/Groups/MDM-Devices-Admin.json",
    "/Groups/MDM-Devices-Android.json",
    "/Groups/MDM-Devices-IOS.json",
  ];

  const groups = await graphClient.api("/groups").get();
  groups.value.forEach((objitem) => {
    if (objitem.displayName === "MDM-Users-Intune") {
      groupid = objitem.id;
    }
  });
  //console.log(groupid);

  const configurationprofileid = [];
  const configurationprofiles = await graphClient
    .api("/deviceManagement/deviceConfigurations")
    .get();

  configurationprofiles.value.forEach((CPofile) => {
    configurationprofileid.push(CPofile.id);
  });

  const configurationprofileAssignments = `{
    "assignments": [
      {
          "target": {
              "@odata.type": "#microsoft.graph.groupAssignmentTarget",
              "groupId": "${groupid}"
          }
      }
  ]
  }`;
  let isAssigned = false;

  function handleClickAssign() {
    if (isAssigned) {
      console.log("Already imported");
      ClickAssign.onclick = null;
      // Import has already been performed, so exit the function
      return;
    }
    configurationprofileid.forEach((CPofileID) => {
      const CProfileAssignments = graphClient
        .api(
          "/deviceManagement/deviceConfigurations" +
            "/" +
            CPofileID +
            "/" +
            "microsoft.graph.assign"
        )
        .version("beta")
        .post(JSON.parse(configurationprofileAssignments));
    });
    console.log("Policy's Assigned");
    isAssigned = true;
  }

  const ClickAssign = document.getElementById("Assign");
  ClickAssign.onclick = handleClickAssign;

  // Array for configurationPolicys
  const deviceManagementConfigurationPolicys = [
    "/SettingsCatalog/Dustin - Default Applications.json",
    //"/SettingsCatalog/Dustin - Default Delivery Optimization.json",
    //"/SettingsCatalog/Dustin - Default Device policy for Windows 10.json",
    //"/SettingsCatalog/Dustin - Diagnostic.json",
  ];
  // Imports ConfigurationPolicys above when clicking "Import"
  let isImported = false;

  function handleClickImport() {
    if (isImported) {
      console.log("Already imported");
      ClickImport.onclick = null;
      // Import has already been performed, so exit the function
      return;
    }
    ImportGroups.forEach((ImportedGroup) => {
      fetch(ImportedGroup)
        .then((response) => response.json())
        .then((importGroup) => {
          graphClient
            .api("/groups")
            .version("beta")
            .post(importGroup)
            .then((response) => {
              // Handle the response
              console.log(response);
            })
            .catch((error) => {
              // Handle the error
              console.error(error);
            });
        })
        .catch((error) => {
          // Handle the error
          console.error(error);
        });
    });

    deviceManagementConfigurationPolicys.forEach(
      (deviceManagementConfigurationPolicyUrl) => {
        fetch(deviceManagementConfigurationPolicyUrl)
          .then((response) => response.json())
          .then((deviceManagementConfigurationPolicy) => {
            // Check if the policy already exists
            graphClient
              .api("/deviceManagement/configurationPolicies")
              .version("beta")
              .filter(`name eq '${deviceManagementConfigurationPolicy.name}'`)
              .get()
              .then((response) => {
                const existingPolicies = response.value;
                if (existingPolicies.length > 0) {
                  console.log(
                    `Policy "${deviceManagementConfigurationPolicy.name}" already exists.`
                  );
                } else {
                  // Create the policy
                  graphClient
                    .api("/deviceManagement/configurationPolicies")
                    .version("beta")
                    .post(deviceManagementConfigurationPolicy)
                    .then((response) => {
                      // Handle the response
                      console.log(response);
                    })
                    .catch((error) => {
                      // Handle the error
                      console.error(error);
                    });
                }
              })
              .catch((error) => {
                // Handle the error
                console.error(error);
              });
          })
          .catch((error) => {
            // Handle the error
            console.error(error);
          });
      }
    );

    // Set the flag to true to indicate that the import has been performed
    isImported = true;
  }

  // Listens to when import is clicked and runs the handleClickImport function
  const ClickImport = document.getElementById("Import");
  ClickImport.onclick = handleClickImport;

  // GET API request for scripts ID
  const deviceManagementScripts = await graphClient
    .api("/deviceManagement/deviceManagementScripts")
    .version("beta")
    .get();
  //console.dir(deviceManagementScripts.value);

  // loop through's each script and gets the ID
  deviceManagementScripts.value.forEach((obj) => {
    const deviceManagementScriptsID = graphClient
      .api("/deviceManagement/deviceManagementScripts" + "/" + obj.id)
      .version("beta")
      .get();
    // Items list begins with array outside of the function
    const items = [];
    // Decodes from base64string
    deviceManagementScriptsID.then(function (result) {
      var encodedStringAtoB = result.scriptContent;
      var decodedStringAtoB = atob(encodedStringAtoB);
      //console.log(result.scriptContent);
      //console.log(decodedStringAtoB);
      // pushes displaynames to the items array above
      items.push(result.displayName);
      // List powershell script
      let isDownloaded = false;

      function createList() {
        const listContainer = document.getElementById("listContainer");

        items.forEach(function (item) {
          const listItem = document.createElement("li");
          const downloadButton = document.createElement("button");
          downloadButton.setAttribute("id", "listbutton");
          downloadButton.textContent = "Download";
          listItem.style.listStyleType = "none";
          downloadButton.onclick = handleClickDownload;

          function handleClickDownload() {
            if (isDownloaded) {
              // Download has already been performed, so exit the function
              return;
            }

            //console.log("Button clicked");
            downloadFile();
            //console.log(item);
          }

          function downloadFile() {
            // File download logic here
            const content = decodedStringAtoB;
            const filename = item + ".ps1";
            const blob = new Blob([content], { type: "text/plain" });

            const downloadLink = document.createElement("a");
            downloadLink.href = URL.createObjectURL(blob);
            downloadLink.download = filename;
            downloadLink.style.display = "none";

            // Append the link to the document body
            document.body.appendChild(downloadLink);

            // Programmatically trigger the download
            downloadLink.click();

            // Cleanup
            document.body.removeChild(downloadLink);

            // Set the flag to true to indicate that the download has been performed
            isDownloaded = true;
          }

          const label = document.createElement("label");
          label.textContent = item;
          listItem.appendChild(downloadButton);
          listItem.appendChild(label);

          listContainer.appendChild(listItem);
        });
      }

      createList();

      let isContentDisplayed = false;

      function displayContent() {
        if (isContentDisplayed) {
          // Content has already been displayed, so exit the function
          return;
        }
        // Display the content (replace this with your actual content display logic)
        console.log("You are already connected");

        // Set the flag to true to indicate that the content has been displayed
        isContentDisplayed = true;
      }

      function handleClickConnect() {
        displayContent();
      }
      const ClickConnect = document.getElementById("Connect");
      ClickConnect.onclick = handleClickConnect;
    });
  });
}
console.log("Test3");
