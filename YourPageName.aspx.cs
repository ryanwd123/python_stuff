using System;
using System.IO;
using System.Web.UI;

namespace YourNamespace
{
    public partial class YourPageName : Page
    {
        protected void btnCreateFile_Click(object sender, EventArgs e)
        {
            try
            {
                // Define the file path and name
                string directoryPath = Server.MapPath("~/YourDirectoryPath/");
                string fileName = DateTime.Now.ToString("yyyy-MM-dd") + ".txt";
                string filePath = Path.Combine(directoryPath, fileName);

                // Create the directory if it doesn't exist
                if (!Directory.Exists(directoryPath))
                {
                    Directory.CreateDirectory(directoryPath);
                }

                // Write "hello world" to the file
                File.WriteAllText(filePath, "hello world");

                // Provide feedback to the user (optional)
                ClientScript.RegisterStartupScript(this.GetType(), "alert", "alert('File created successfully!');", true);
            }
            catch (Exception ex)
            {
                // Handle exceptions (optional)
                ClientScript.RegisterStartupScript(this.GetType(), "alert", "alert('An error occurred: " + ex.Message + "');", true);
            }
        }
    }
}
