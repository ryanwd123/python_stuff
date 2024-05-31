<%@ Page Language="C#" %>
<!DOCTYPE html>
<html>
<head>
    <title>Create File in SharePoint Library</title>
    <script src="https://code.jquery.com/jquery-3.6.0.min.js"></script>
    <script type="text/javascript">
        $(document).ready(function() {
            $('#createFileButton').click(function() {
                createTextFile();
            });
        });

        function createTextFile() {
            var siteUrl = _spPageContextInfo.webAbsoluteUrl; // Get the current site URL
            var folderUrl = "/Shared Documents"; // Path to the document library
            var fileName = "NewFile.txt"; // Name of the new file
            var fileContent = "This is the content of the new text file."; // Content of the new file

            // Build the REST API endpoint URL
            var apiUrl = siteUrl + "/_api/web/GetFolderByServerRelativeUrl('" + folderUrl + "')/Files/add(url='" + fileName + "', overwrite=true)";

            $.ajax({
                url: apiUrl,
                type: "POST",
                data: fileContent,
                headers: {
                    "accept": "application/json;odata=verbose",
                    "X-RequestDigest": $("#__REQUESTDIGEST").val(),
                    "content-type": "text/plain; charset=utf-8"
                },
                success: function(data) {
                    alert("File created successfully!");
                },
                error: function(error) {
                    alert("Error: " + JSON.stringify(error));
                }
            });
        }
    </script>
</head>
<body>
    <form id="aspnetForm" runat="server">
        <div>
            <asp:ScriptManager runat="server" />
            <asp:Button ID="createFileButton" runat="server" Text="Create Text File" OnClientClick="createTextFile(); return false;" />
        </div>
    </form>
</body>
</html>
