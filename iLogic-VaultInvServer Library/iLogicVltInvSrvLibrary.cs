using System;
using System.Collections.Generic;
using System.Linq;
using System.Drawing;

using ACW = Autodesk.Connectivity.WebServices;
using AWT = Autodesk.Connectivity.WebServicesTools;
using VDF = Autodesk.DataManagement.Client.Framework;


namespace QuickstartiLogicVltInvSrvLibrary
{
    /// <summary>
    /// Collection of functions querying and downloading Vault files for iLogic on Inventor Server.
    /// Note - this collection is a subset of the Quickstart iLogic-Vault Library and includes only methods to download files;
    /// methods to create new files are excluded because iLogic for Inventor Server is not entitled to check-in new files.
    /// </summary>
    public class iLogicVltInvSrvLibrary : IDisposable
    {
        /// <summary>
        /// Empty function, prepared to dispose data if future additions require.
        /// </summary>
        public void Dispose()
        {
            //do clean up here if required
        }

        /// <summary>
        /// Any Vault interaction requires an active Client-Server connection.
        /// </summary>
        private VDF.Vault.Currency.Connections.Connection conn;

        /// <summary>
        /// For VaultInventorServer application only: Re-uses the job processors log-in user Id and ticket for iLogic-Vault interactions within rules.
        /// </summary>
        /// <param name="DbSrvName"></param>
        /// <param name="FlSrvName"></param>
        /// <param name="VaultName"></param>
        /// <param name="UserId"></param>
        /// <param name="Ticket"></param>
        /// <returns>Returns true, if connection is valid</returns>
        public bool ReuseConnection(string DbSrvName, string FlSrvName, string VaultName, long UserId, string Ticket)
        {
            ACW.ServerIdentities mSrvIdnts = new ACW.ServerIdentities();
            mSrvIdnts.DataServer = DbSrvName;
            mSrvIdnts.FileServer = FlSrvName;
            AWT.UserIdTicketCredentials mCred = new AWT.UserIdTicketCredentials(mSrvIdnts, VaultName, UserId, Ticket);
            Autodesk.Connectivity.WebServicesTools.WebServiceManager mWsMgr = new AWT.WebServiceManager(mCred);
            VDF.Vault.Currency.Connections.Connection mConnection = new VDF.Vault.Currency.Connections.Connection(
                mWsMgr, VaultName, UserId, DbSrvName, VDF.Vault.Currency.Connections.AuthenticationFlags.Standard);
            if (mConnection != null)
            {
                conn = mConnection;
                return true;
            }
            return false;
        }


        /// <summary>
        /// Adds local file to Vault.
        /// </summary>
        /// <param name="FullFileName">File path and name of file to add in local working folder.</param>
        /// <param name="VaultFolderPath">Full path in Vault, e.g. "$/Designs/P-00000</param>
        /// <param name="UpdateExisting">Creates new file version if existing file is available for check-out.</param>
        /// <returns>Returns True/False on success/failure; returns false if the file exists and UpdateExisting = false. Returns false for IAM, IPN, IDW/DWG</returns>
        public bool AddFile(string FullFileName, string VaultFolderPath, bool UpdateExisting = true)
        {
            Autodesk.Connectivity.WebServicesTools.WebServiceManager mWsMgr = conn.WebServiceManager;
            System.IO.FileInfo mLocalFileInfo = new System.IO.FileInfo(FullFileName);

            ACW.Folder mFolder = mWsMgr.DocumentService.FindFoldersByPaths(new string[] { VaultFolderPath }).FirstOrDefault();
            if (mFolder == null || mFolder.Id == -1)
            {
                return false;
            }
            string vaultFilePath = System.IO.Path.Combine(mFolder.FullName, mLocalFileInfo.Name).Replace("\\", "/");

            ACW.File wsFile = mWsMgr.DocumentService.FindLatestFilesByPaths(new string[] { vaultFilePath }).First();

            //exclude CAD files with references
            if (IsCadFile(wsFile, mWsMgr))
            {
                return false;
            }

            VDF.Currency.FilePathAbsolute vdfPath = new VDF.Currency.FilePathAbsolute(mLocalFileInfo.FullName);
            VDF.Vault.Currency.Entities.FileIteration vdfFile = null;
            VDF.Vault.Currency.Entities.FileIteration addedFile = null;
            VDF.Vault.Currency.Entities.FileIteration mUploadedFile = null;
            if (wsFile == null || wsFile.Id < 0)
            {
                // add new file to Vault
                var folderEntity = new Autodesk.DataManagement.Client.Framework.Vault.Currency.Entities.Folder(conn, mFolder);
                try
                {
                    addedFile = conn.FileManager.AddFile(folderEntity, "Added by iLogic rule", null, null, ACW.FileClassification.None, false, vdfPath);
                    return true;
                }
                catch (Exception ex)
                {
                    throw new Exception("iLogic rule could not add file " + vdfPath + "Exception: ", ex);
                }

            }
            else
            {
                if (UpdateExisting == true)
                {
                    // checkin new file version
                    VDF.Vault.Settings.AcquireFilesSettings aqSettings = new VDF.Vault.Settings.AcquireFilesSettings(conn)
                    {
                        DefaultAcquisitionOption = VDF.Vault.Settings.AcquireFilesSettings.AcquisitionOption.Checkout
                    };
                    vdfFile = new VDF.Vault.Currency.Entities.FileIteration(conn, wsFile);
                    aqSettings.AddEntityToAcquire(vdfFile);
                    var results = conn.FileManager.AcquireFiles(aqSettings);
                    try
                    {
                        mUploadedFile = conn.FileManager.CheckinFile(results.FileResults.First().File, "Created by iLogic rule", false, null, null, false, null, ACW.FileClassification.None, false, vdfPath);
                        return true;
                    }
                    catch (Exception ex)
                    {
                        throw new Exception("iLogic rule could not update existing file " + vdfFile + "Exception: ", ex);
                    }
                }
                else
                {
                    return false;
                }
            }
        }

        private bool IsCadFile(ACW.File wsFile, Autodesk.Connectivity.WebServicesTools.WebServiceManager mWsMgr)
        {
            //don't add Inventor files except single part files
            List<string> mFileExtensions = new List<string> { ".iam", ".idw", ".dwg" };
            if (mFileExtensions.Any(n => wsFile.Name.EndsWith(n)))
            {
                return true;
            }
            //differentiate Inventor DWG and AutoCAD DWG
            if (wsFile.Name.EndsWith(".dwg"))
            {
                ACW.PropDef[] mPropDefs = mWsMgr.PropertyService.GetPropertyDefinitionsByEntityClassId("FILE");
                ACW.PropDef mPropDef = null;
                ACW.PropInst mPropInst = null;
                mPropDef = mPropDefs.Where(n => n.SysName == "Provider").FirstOrDefault();
                mPropInst = (mWsMgr.PropertyService.GetPropertiesByEntityIds("FILE", new long[] { wsFile.Id })).Where(n => n.PropDefId == mPropDef.Id).FirstOrDefault();
                if (mPropInst.Val.ToString() == "Inventor DWG")
                {
                    return true;
                }
            }
            return false;
        }


        /// <summary>
        /// 
        /// </summary>
        /// <param name="LocalPath">File or Folder path in local working folder</param>
        /// <returns>Vault Folder Path; if LocalPath is a Filepath, the file's parent Folderpath returns</returns>
        public string ConvertLocalPathToVaultPath(string LocalPath)
        {
            string mVaultPath = null;
            string mWf = conn.WorkingFoldersManager.GetWorkingFolder("$/").FullPath;
            if (LocalPath.Contains(mWf))
            {
                if (IsFilePath(LocalPath) == true)
                {
                    System.IO.FileInfo fileInfo = new System.IO.FileInfo(LocalPath);
                    LocalPath = fileInfo.DirectoryName;
                }
                if (IsDirPath(LocalPath) == true)
                {
                    mVaultPath = LocalPath.Replace(mWf, "$/");
                    mVaultPath = mVaultPath.Replace("\\", "/");
                    return mVaultPath;
                }
                else
                {
                    return "Invalid local path";
                }
            }
            else
            {
                return "Error: Local path outside of working folder";
            }
        }


        private bool IsFilePath(string path)
        {
            if (System.IO.File.Exists(path)) return true;
            return false;
        }

        private bool IsDirPath(string path)
        {
            if (System.IO.Directory.Exists(path)) return true;
            return false;
        }

        /// <summary>
        /// Downloads Vault file using full file path, e.g. "$/Designs/Base.ipt". Returns full file name in local working folder (download enforces override, if local file exists),
        /// returns "FileNotFound if file does not exist at indicated location.
        /// Preset Options: Download Children (recursively) = Enabled, Enforce Overwrite = True
        /// </summary>
        /// <param name="VaultFullFileName">Full Vault File Path of format '$\...\*.*'</param>
        /// <param name="CheckOut">Optional. File downloaded does NOT check-out as default.</param>
        /// <returns>Local path/filename</returns>
        public string GetFileByFullFilePath(string VaultFullFileName, bool CheckOut = false)
        {
            List<string> mFiles = new List<string>();
            List<string> mFilesDownloaded = new List<string>();
            mFiles.Add(VaultFullFileName);
            ACW.File[] wsFiles = conn.WebServiceManager.DocumentService.FindLatestFilesByPaths(mFiles.ToArray());
            VDF.Vault.Currency.Entities.FileIteration mFileIt = new VDF.Vault.Currency.Entities.FileIteration(conn, (wsFiles[0]));

            //define download options, including DefaultAcquisitionOptions
            VDF.Vault.Settings.AcquireFilesSettings settings = CreateAcquireSettings(false);
            settings.AddFileToAcquire(mFileIt, settings.DefaultAcquisitionOption);

            //download
            VDF.Vault.Results.AcquireFilesResults results = conn.FileManager.AcquireFiles(settings);

            if (CheckOut)
            {
                //define checkout options and checkout
                settings = CreateAcquireSettings(true);
                settings.AddFileToAcquire(mFileIt, settings.DefaultAcquisitionOption);
                results = conn.FileManager.AcquireFiles(settings);
            }

            //refine and validate output
            if (results != null)
            {
                try
                {
                    if (results.FileResults.Any(n => n.File.EntityName == mFileIt.EntityName))
                    {
                        mFilesDownloaded.Add(conn.WorkingFoldersManager.GetPathOfFileInWorkingFolder(mFileIt).FullPath.ToString());
                    }

                    return mFilesDownloaded[0];

                }
                catch (Exception)
                {
                    return null;
                }
            }
            return null;
        }

        /// <summary>
        /// Find 1 to many file(s) by 1 to many search criteria as property/value pairs. 
        /// Downloads first file matching all or any search criterias. 
        /// Preset Search Operator Options: [Property] is (exactly) [Value]; multiple conditions link up using AND condition.
        /// Preset Download Options: Download Children (recursively) = Enabled, Enforce Overwrite = True
        /// </summary>
        /// <param name="SearchCriteria">Dictionary of property/value pairs</param>
        /// <param name="MatchAllCriteria">Optional. Switches AND/OR conditions using multiple criterias. Default is true</param>
        /// <param name="CheckOut">Optional. File downloaded does NOT check-out as default</param>
        /// <param name="FoldersSearched">Optional. Limit search scope to given folder path(s).</param>
        /// <returns>Local path/filename</returns>
        public string GetFileBySearchCriteria(Dictionary<string, string> SearchCriteria, bool MatchAllCriteria = true, bool CheckOut = false, string[] FoldersSearched = null)
        {
            //FoldersSearched: Inventor files are expected in IPJ registered path's only. In case of null use these:
            ACW.Folder[] mFldr;
            List<long> mFolders = new List<long>();
            if (FoldersSearched != null)
            {
                mFldr = conn.WebServiceManager.DocumentService.FindFoldersByPaths(FoldersSearched);
                foreach (ACW.Folder folder in mFldr)
                {
                    if (folder.Id != -1) mFolders.Add(folder.Id);
                }
            }

            List<String> mFilesFound = new List<string>();
            List<String> mFilesDownloaded = new List<string>();
            //combine all search criteria
            List<ACW.SrchCond> mSrchConds = CreateSrchConds(SearchCriteria, MatchAllCriteria);
            List<ACW.File> totalResults = new List<ACW.File>();
            string bookmark = string.Empty;
            ACW.SrchStatus status = null;

            while (status == null || totalResults.Count < status.TotalHits)
            {
                ACW.File[] mSrchResults = conn.WebServiceManager.DocumentService.FindFilesBySearchConditions(
                    mSrchConds.ToArray(), null, mFolders.ToArray(), true, true, ref bookmark, out status);
                if (mSrchResults != null) totalResults.AddRange(mSrchResults);
                else break;
            }
            //if results not empty
            if (totalResults.Count >= 1)
            {
                ACW.File wsFile = totalResults.First<ACW.File>();
                VDF.Vault.Currency.Entities.FileIteration mFileIt = new VDF.Vault.Currency.Entities.FileIteration(conn, (wsFile));

                //build download options including DefaultAcquisitionOptions
                VDF.Vault.Settings.AcquireFilesSettings settings = CreateAcquireSettings(false);
                settings.AddFileToAcquire(mFileIt, settings.DefaultAcquisitionOption);

                //download
                VDF.Vault.Results.AcquireFilesResults results = conn.FileManager.AcquireFiles(settings);

                if (CheckOut)
                {
                    //define checkout options and checkout
                    settings = CreateAcquireSettings(true);
                    settings.AddFileToAcquire(mFileIt, settings.DefaultAcquisitionOption);
                    results = conn.FileManager.AcquireFiles(settings);
                }

                //refine and validate output
                if (results != null)
                {
                    try
                    {
                        if (results.FileResults.Any(n => n.File.EntityName == mFileIt.EntityName))
                        {
                            mFilesDownloaded.Add(conn.WorkingFoldersManager.GetPathOfFileInWorkingFolder(mFileIt).FullPath.ToString());
                        }

                        return mFilesDownloaded[0];

                    }
                    catch (Exception)
                    {
                        return null;
                    }
                }
                else
                {
                    return null;
                }
            }
            else
            {
                return null;
            }
        }

        /// <summary>
        /// Find 1 to many file(s) by 1 to many search criteria as property/value pairs. 
        /// Downloads all files found, matching the criteria. Returns array of full file names of downloaded files
        /// Preset Search Operator Options: [Property] is (exactly) [Value]; multiple conditions link up using AND/OR condition, depending MatchAllCriteria = True/False
        /// </summary>
        /// <param name="SearchCriteria">Dictionary of property/value pairs</param>
        /// <param name="MatchAllCriteria">Optional. Switches AND/OR conditions using multiple criterias. Default is true</param>
        /// <param name="CheckOut">Optional. Downloaded files will NOT check-out as default.</param>
        /// <param name="FoldersSearched">Optional. Limit search scope to given folder path(s).</param>
        /// <returns>Array of file names found</returns>
        public IList<string> GetFilesBySearchCriteria(Dictionary<string, string> SearchCriteria, bool MatchAllCriteria = true, bool CheckOut = false, string[] FoldersSearched = null)
        {
            //FoldersSearched: Inventor files are expected in IPJ registered path's only. In case of null use these:
            ACW.Folder[] mFldr;
            List<long> mFolders = new List<long>();
            if (FoldersSearched != null)
            {
                mFldr = conn.WebServiceManager.DocumentService.FindFoldersByPaths(FoldersSearched);
                foreach (ACW.Folder folder in mFldr)
                {
                    if (folder.Id != -1) mFolders.Add(folder.Id);
                }
            }

            List<VDF.Vault.Currency.Entities.FileIteration> mFilesFound = new List<VDF.Vault.Currency.Entities.FileIteration>();
            List<String> mFilesDownloaded = new List<string>();
            //combine all search criteria
            List<ACW.SrchCond> mSrchConds = CreateSrchConds(SearchCriteria, MatchAllCriteria);
            List<ACW.File> totalResults = new List<ACW.File>();
            string bookmark = string.Empty;
            ACW.SrchStatus status = null;

            while (status == null || totalResults.Count < status.TotalHits)
            {
                ACW.File[] mSrchResults = conn.WebServiceManager.DocumentService.FindFilesBySearchConditions(
                    mSrchConds.ToArray(), null, mFolders.ToArray(), true, true, ref bookmark, out status);
                if (mSrchResults != null) totalResults.AddRange(mSrchResults);
                else break;
            }
            //if results not empty
            if (totalResults.Count >= 1)
            {
                //build download options including DefaultAcquisitionOptions
                VDF.Vault.Settings.AcquireFilesSettings settings = CreateAcquireSettings(false);
                foreach (ACW.File wsFile in totalResults)
                {
                    VDF.Vault.Currency.Entities.FileIteration mFileIt = new VDF.Vault.Currency.Entities.FileIteration(conn, wsFile);
                    //register file paths to validate downloaded ones later
                    mFilesFound.Add(mFileIt);
                    settings.AddFileToAcquire(mFileIt, settings.DefaultAcquisitionOption);
                }

                //download
                VDF.Vault.Results.AcquireFilesResults results = conn.FileManager.AcquireFiles(settings);

                if (CheckOut)
                {
                    //define checkout options and checkout
                    settings = CreateAcquireSettings(true);
                    foreach (ACW.File wsFile in totalResults)
                    {
                        VDF.Vault.Currency.Entities.FileIteration mFileIt = new VDF.Vault.Currency.Entities.FileIteration(conn, wsFile);
                        settings.AddFileToAcquire(mFileIt, settings.DefaultAcquisitionOption);
                    }

                    //checkout
                    results = conn.FileManager.AcquireFiles(settings);
                }

                //refine and validate output
                if (results.FileResults != null)
                {
                    foreach (VDF.Vault.Currency.Entities.FileIteration mFileIt in mFilesFound)
                    {
                        if (results.FileResults.Any(n => n.File.EntityName == mFileIt.EntityName))
                        {
                            mFilesDownloaded.Add(conn.WorkingFoldersManager.GetPathOfFileInWorkingFolder(mFileIt).FullPath.ToString());
                        }
                    }
                    return mFilesDownloaded;
                }
                return null;
            }
            else
            {
                return null;
            }
        }

        /// <summary>
        /// Download Thumbnail Image of the given file as Image file.
        /// </summary>
        /// <param name="VaultFullFileName">Full Vault source file path of format '$\...\*.*'</param>
        /// <param name="Width">Optional. Image pixel size</param>
        /// <param name="Height">Optional. Image pixel size.</param>
        /// <returns>Full file path of image file (*.jpg)</returns>
        public string GetThumbnailFileByFullSourceFilePath(string VaultFullFileName, int Width = 300, int Height = 300)
        {
            string mImageFullFileName = null;
            List<string> mFiles = new List<string>();
            mFiles.Add(VaultFullFileName);

            try
            {
                ACW.File[] wsFiles = conn.WebServiceManager.DocumentService.FindLatestFilesByPaths(mFiles.ToArray());
                VDF.Vault.Currency.Entities.FileIteration mFileIt = new VDF.Vault.Currency.Entities.FileIteration(conn, (wsFiles[0]));

                Image image = GetThumbnailImage(mFileIt, Width, Height);
                if (image != null)
                {
                    string mExt = System.IO.Path.GetExtension(VaultFullFileName);
                    string LocalFullFileName = conn.WorkingFoldersManager.GetPathOfFileInWorkingFolder(mFileIt).ToString();
                    mImageFullFileName = LocalFullFileName.Replace(mExt, ".jpg");
                    image.Save(mImageFullFileName);
                }
                System.IO.FileInfo fileInfo = new System.IO.FileInfo(mImageFullFileName);
                if (fileInfo.Exists)
                {
                    return mImageFullFileName;
                }
                else { return null; }
            }
            catch (Exception)
            {
                return null;
            }
        }


        /// <summary>
        /// Download Thumbnail Image of the file searched as Image file.
        /// </summary>
        /// <param name="SearchCriteria">Dictionary of property/value pairs</param>
        /// <param name="MatchAllCriteria">Optional. Switches AND/OR conditions using multiple criterias. Default is true</param>
        /// <param name="FoldersSearched">Optional. Limit search scope to given folder path(s).</param>
        /// <param name="Width">Optional. Image pixel size</param>
        /// <param name="Height">Optional. Image pixel size.</param>
        /// <returns>Full file path of image file (*.jpg)</returns>
        public string GetThumbnailFileBySearchCriteria(Dictionary<string, string> SearchCriteria, bool MatchAllCriteria = true, string[] FoldersSearched = null, int Width = 300, int Height = 300)
        {
            //FoldersSearched: Inventor files are expected in IPJ registered path's only. In case of null use these:
            ACW.Folder[] mFldr;
            List<long> mFolders = new List<long>();
            if (FoldersSearched != null)
            {
                mFldr = conn.WebServiceManager.DocumentService.FindFoldersByPaths(FoldersSearched);
                foreach (ACW.Folder folder in mFldr)
                {
                    if (folder.Id != -1) mFolders.Add(folder.Id);
                }
            }

            List<String> mFilesFound = new List<string>();
            //combine all search criteria
            List<ACW.SrchCond> mSrchConds = CreateSrchConds(SearchCriteria, MatchAllCriteria);
            List<ACW.File> totalResults = new List<ACW.File>();
            string bookmark = string.Empty;
            ACW.SrchStatus status = null;

            while (status == null || totalResults.Count < status.TotalHits)
            {
                ACW.File[] mSrchResults = conn.WebServiceManager.DocumentService.FindFilesBySearchConditions(
                    mSrchConds.ToArray(), null, mFolders.ToArray(), true, true, ref bookmark, out status);
                if (mSrchResults != null) totalResults.AddRange(mSrchResults);
                else break;
            }
            //if results not empty
            if (totalResults.Count >= 1)
            {
                try
                {
                    string mImageFullFileName = null;
                    ACW.File wsFile = totalResults.First<ACW.File>();
                    VDF.Vault.Currency.Entities.FileIteration mFileIt = new VDF.Vault.Currency.Entities.FileIteration(conn, (wsFile));

                    Image image = GetThumbnailImage(mFileIt, Height, Width);
                    if (image != null)
                    {
                        ACW.Folder mParentFldr = conn.WebServiceManager.DocumentService.GetFolderById(wsFile.FolderId);
                        string VaultFullFileName = mParentFldr.FullName + "/" + wsFile.Name;
                        string mExt = System.IO.Path.GetExtension(VaultFullFileName);
                        string LocalFullFileName = conn.WorkingFoldersManager.GetPathOfFileInWorkingFolder(mFileIt).ToString();
                        mImageFullFileName = LocalFullFileName.Replace(mExt, ".jpg");
                        image.Save(mImageFullFileName);
                    }
                    System.IO.FileInfo fileInfo = new System.IO.FileInfo(mImageFullFileName);
                    if (fileInfo.Exists)
                    {
                        return mImageFullFileName;
                    }
                    else { return null; }
                }
                catch (Exception)
                {
                    return null;
                }
            }
            return null;
        }


        private VDF.Vault.Settings.AcquireFilesSettings CreateAcquireSettings(bool CheckOut)
        {
            VDF.Vault.Settings.AcquireFilesSettings settings = new VDF.Vault.Settings.AcquireFilesSettings(conn);
            if (CheckOut)
            {
                settings.DefaultAcquisitionOption = VDF.Vault.Settings.AcquireFilesSettings.AcquisitionOption.Checkout;
            }
            else
            {
                settings.DefaultAcquisitionOption = VDF.Vault.Settings.AcquireFilesSettings.AcquisitionOption.Download;
                settings.OptionsRelationshipGathering.FileRelationshipSettings.IncludeChildren = true;
                settings.OptionsRelationshipGathering.FileRelationshipSettings.RecurseChildren = true;
                settings.OptionsRelationshipGathering.FileRelationshipSettings.IncludeAttachments = true;
                settings.OptionsRelationshipGathering.FileRelationshipSettings.IncludeLibraryContents = true;
                settings.OptionsRelationshipGathering.FileRelationshipSettings.ReleaseBiased = true;
                settings.OptionsRelationshipGathering.FileRelationshipSettings.VersionGatheringOption = VDF.Vault.Currency.VersionGatheringOption.Revision;
                settings.OptionsRelationshipGathering.IncludeLinksSettings.IncludeLinks = false;
                VDF.Vault.Settings.AcquireFilesSettings.AcquireFileResolutionOptions mResOpt = new VDF.Vault.Settings.AcquireFilesSettings.AcquireFileResolutionOptions();
                mResOpt.OverwriteOption = VDF.Vault.Settings.AcquireFilesSettings.AcquireFileResolutionOptions.OverwriteOptions.ForceOverwriteAll;
                mResOpt.SyncWithRemoteSiteSetting = VDF.Vault.Settings.AcquireFilesSettings.SyncWithRemoteSite.Always;
            }

            return settings;
        }

        private List<ACW.SrchCond> CreateSrchConds(Dictionary<string, string> SearchCriteria, bool MatchAllCriteria)
        {
            ACW.PropDef[] mFilePropDefs = conn.WebServiceManager.PropertyService.GetPropertyDefinitionsByEntityClassId("FILE");
            //iterate mSearchcriteria to get property definitions and build AWS search criteria
            List<ACW.SrchCond> mSrchConds = new List<ACW.SrchCond>();
            int i = 0;
            foreach (var item in SearchCriteria)
            {
                ACW.PropDef mFilePropDef = mFilePropDefs.Single(n => n.DispName == item.Key);
                ACW.SrchCond mSearchCond = new ACW.SrchCond();
                {
                    mSearchCond.PropDefId = mFilePropDef.Id;
                    mSearchCond.PropTyp = ACW.PropertySearchType.SingleProperty;
                    mSearchCond.SrchOper = 3; //equals
                    if (MatchAllCriteria) mSearchCond.SrchRule = ACW.SearchRuleType.Must;
                    else mSearchCond.SrchRule = ACW.SearchRuleType.May;
                    mSearchCond.SrchTxt = item.Value;
                }
                mSrchConds.Add(mSearchCond);
                i++;
            }
            return mSrchConds;
        }

        private Image GetThumbnailImage(VDF.Vault.Currency.Entities.FileIteration fileIteration, int Width, int Height)
        {
            try
            {
                VDF.Vault.Currency.Properties.PropertyDefinitionDictionary mPropDefs = conn.PropertyManager.GetPropertyDefinitions(
                                        VDF.Vault.Currency.Entities.EntityClassIds.Files, null, VDF.Vault.Currency.Properties.PropertyDefinitionFilter.IncludeSystem);
                VDF.Vault.Currency.Properties.PropertyDefinition mThmbNlPropDef = (mPropDefs.SingleOrDefault(n => n.Key == "Thumbnail").Value);
                VDF.Vault.Currency.Properties.PropertyValueSettings mPropSetting = new VDF.Vault.Currency.Properties.PropertyValueSettings();

                VDF.Vault.Currency.Properties.ThumbnailInfo mThumbInfo = (VDF.Vault.Currency.Properties.ThumbnailInfo)conn.PropertyManager.GetPropertyValue(fileIteration, mThmbNlPropDef, mPropSetting);

                Image image = RenderThumbnailToImage(mThumbInfo, Height, Width);
                return image;
            }
            catch
            {
                return null;
            }
        }

        /// <summary>
        /// Vault Blog Sample function to convert legacy meta file format and image file format (added to Vault 2013 and later)
        /// </summary>
        /// <param name="value">Vault Image property types return ThumbnailInfo</param>
        /// <param name="width">recommended max. size = 300, but custom thumbnails may be larger</param>
        /// <param name="height">recommended max. size = 300, but custom thumbnails may be larger</param>
        /// <returns></returns>
        private static Image RenderThumbnailToImage(VDF.Vault.Currency.Properties.ThumbnailInfo value, int width, int height)
        {
            // convert the property value to a byte array
            byte[] thumbnailRaw = value.Image as byte[];

            if (null == thumbnailRaw || 0 == thumbnailRaw.Length)
                return null;

            Image retImage = null;

            using (System.IO.MemoryStream memStream = new System.IO.MemoryStream(thumbnailRaw))
            {
                using (System.IO.BinaryReader br = new System.IO.BinaryReader(memStream))
                {
                    int CF_METAFILEPICT = 3;
                    int CF_ENHMETAFILE = 14;

                    int clipboardFormatId = br.ReadInt32(); /*int clipFormat =*/
                    bool bytesRepresentMetafile = (clipboardFormatId == CF_METAFILEPICT || clipboardFormatId == CF_ENHMETAFILE);
                    try
                    {

                        if (bytesRepresentMetafile)
                        {
                            // the bytes represent a clipboard metafile

                            // read past header information
                            br.ReadInt16();
                            br.ReadInt16();
                            br.ReadInt16();
                            br.ReadInt16();

                            System.Drawing.Imaging.Metafile mf = new System.Drawing.Imaging.Metafile(br.BaseStream);
                            retImage = mf.GetThumbnailImage(width, height, new Image.GetThumbnailImageAbort(GetThumbnailImageAbort), IntPtr.Zero);
                        }
                        else
                        {
                            // the bytes do not represent a metafile, try to convert to an Image
                            memStream.Seek(0, System.IO.SeekOrigin.Begin);
                            Image im = Image.FromStream(memStream, true, false);

                            retImage = im.GetThumbnailImage(width, height, new Image.GetThumbnailImageAbort(GetThumbnailImageAbort), IntPtr.Zero);
                        }
                    }
                    catch
                    {
                    }
                }
            }

            return retImage;
        }

        private static bool GetThumbnailImageAbort()
        {
            return false;
        }
    }
}
