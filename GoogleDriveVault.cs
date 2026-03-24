using Google.Apis.Auth.OAuth2;
using Google.Apis.Drive.v3;
using Google.Apis.Services;
using Google.Apis.Util.Store;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq; // Απαραίτητο για το FirstOrDefault()
using System.Threading;
using System.Threading.Tasks;

namespace EuroSearchApp
{
    public class GoogleDriveVault
    {
        static string[] Scopes = { DriveService.Scope.DriveFile };
        static string ApplicationName = "EuroStatement";

        public async Task UploadFileToDrive(string localFilePath)
        {
            try
            {
                UserCredential credential;
                string credPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Assets", "client_secrets.json");

                using (var stream = new FileStream(credPath, FileMode.Open, FileAccess.Read))
                {
                    string tokenPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "token_storage");
                    credential = await GoogleWebAuthorizationBroker.AuthorizeAsync(
                        GoogleClientSecrets.FromStream(stream).Secrets,
                        Scopes,
                        "user",
                        CancellationToken.None,
                        new FileDataStore(tokenPath, true));
                }

                var service = new DriveService(new BaseClientService.Initializer()
                {
                    HttpClientInitializer = credential,
                    ApplicationName = ApplicationName,
                });

                // 1. Βρίσκουμε ή φτιάχνουμε τον κεντρικό φάκελο "EuroStatements"
                string mainFolderId = await GetOrCreateFolder(service, "EuroStatements", null);

                // 2. Δημιουργούμε το όνομα της ημέρας (π.χ. "24 Μαρτίου 2026")
                // Χρησιμοποιούμε "dd MMMM yyyy" για πλήρη ημερομηνία
                string todayFolderName = DateTime.Now.ToString("dd MMMM yyyy", new System.Globalization.CultureInfo("el-GR"));

                // 3. Βρίσκουμε ή φτιάχνουμε τον υποφάκελο της ημέρας ΜΕΣΑ στον κεντρικό φάκελο
                string dailyFolderId = await GetOrCreateFolder(service, todayFolderName, mainFolderId);

                var fileMetadata = new Google.Apis.Drive.v3.Data.File()
                {
                    Name = Path.GetFileName(localFilePath),
                    MimeType = "application/pdf",
                    Parents = new List<string> { dailyFolderId }
                };

                using (var stream = new FileStream(localFilePath, FileMode.Open))
                {
                    var uploadRequest = service.Files.Create(fileMetadata, stream, "application/pdf");
                    uploadRequest.Fields = "id";
                    await uploadRequest.UploadAsync();
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Drive Error: " + ex.Message);
            }
        }

        private async Task<string> GetOrCreateFolder(DriveService service, string folderName, string parentId)
        {
            var request = service.Files.List();

            // Αν έχουμε parentId, ψάχνουμε ΜΕΣΑ σε αυτόν τον φάκελο, αλλιώς στο root
            string query = $"mimeType = 'application/vnd.google-apps.folder' and name = '{folderName}' and trashed = false";
            if (!string.IsNullOrEmpty(parentId))
            {
                query += $" and '{parentId}' in parents";
            }

            request.Q = query;
            var result = await request.ExecuteAsync();
            var folder = result.Files.FirstOrDefault();

            if (folder != null) return folder.Id;

            var newFolder = new Google.Apis.Drive.v3.Data.File()
            {
                Name = folderName,
                MimeType = "application/vnd.google-apps.folder"
            };

            if (!string.IsNullOrEmpty(parentId))
            {
                newFolder.Parents = new List<string> { parentId };
            }

            var createRequest = service.Files.Create(newFolder);
            createRequest.Fields = "id";
            var folderResource = await createRequest.ExecuteAsync();
            return folderResource.Id;
        }
    }
}