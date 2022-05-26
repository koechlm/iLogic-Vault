using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;

using ACET = Autodesk.Connectivity.Explorer.ExtensibilityTools;
using ACW = Autodesk.Connectivity.WebServices;
using VDF = Autodesk.DataManagement.Client.Framework;
using VltBase = Connectivity.Application.VaultBase;
using Autodesk.DataManagement.Client.Framework.Vault.Currency.Properties;

namespace QuickstartiLogicLibrary
{
    /// <summary>
    /// Collection of functions querying and downloading Vault files for iLogic.
    /// </summary>
    public class QuickstartiLogicLib
    {
        /// <summary>
        /// Empty function, prepared to dispose data if future additions require. change class to :IDisposable if needed
        /// </summary>
        private void Dispose()
        {
            //do clean up here if required
        }

        /// <summary>
        /// Any Vault interaction requires an active Client-Server connection.
        /// To avoid Vault API specific references, check connection state using the loggedIn property.
        /// </summary>
        private readonly VDF.Vault.Currency.Connections.Connection conn = VltBase.ConnectionManager.Instance.Connection;

        /// <summary>
        /// Some methods are not applicable to Vault Basic; we need to know the environment
        /// </summary>
        private static readonly ACW.Product[] mProducts = VltBase.ConnectionManager.Instance.Connection.WebServiceManager.InformationService.GetSupportedProducts();

        private readonly bool IsVaultPro = mIsVaultPro;

        private readonly bool IsVaultBasic = mIsVaultBasic;

        /// <summary>
        /// Indicates Vault Basic environment
        /// </summary>
        private static bool mIsVaultBasic
        {
            get
            {
                //Vault Basic Servers return only a single product, whereas Vault Pro returns 4
                if (mProducts.Length == 1)
                {
                    return true;
                }
                return false;
            }
        }

        private static bool mIsVaultPro
        {
            get
            {
                //Vault Pro return 4 products
                if (mProducts.Length == 4)
                {
                    return true;
                }
                return false;
            }
        }

        //avoid multiple server calls for the iLogic-Vault session
        private ACW.PropDef[] propDefs = null;

        /// <summary>
        /// Property representing the current user's Vault connection state; returns true, if current user is logged in.
        /// </summary>
        public bool LoggedIn
        {
            get
            {
                if (conn != null)
                {
                    return true;
                }
                return false;
            }
        }


        /// <summary>
        /// Deprecated. Returns current Vault connection. Leverage LoggedIn property whenever possible. 
        /// Null value returned if user is not logged in.
        /// </summary>
        /// <returns>Vault Connection</returns>
        public VDF.Vault.Currency.Connections.Connection GetVaultConnection()
        {
            if (conn != null)
            {
                return conn;
            }
            return null;
        }

        /// <summary>
        /// Adds local file to Vault and optionally attaches it to a parent file.
        /// </summary>
        /// <param name="FullFileName">File path and name of file to add in local working folder.</param>
        /// <param name="VaultFolderPath">Full path in Vault, e.g. "$/Designs/P-00000</param>
        /// <param name="UpdateExisting">Creates new file version if existing file is available for check-out.</param>
        /// <param name="DisableCheckInDesignFiles">The recommended default blocking Inventor, AutoCAD, Navisworks, Microstation, Solidworks and PRO-E files to maintain file relationships. Disable for exeptions like single dwg export files only.</param>
        /// <param name="ParentFileToAttachTo">Creates an attachment to the parent file consuming the newly added file; 
        /// provide Vault path and file name ('$/...') of parent file to attach to</param>
        /// <returns>Returns True/False on success/failure; returns false if the file exists and UpdateExisting = false. Returns false for IAM, IPN, IDW/DWG.
        /// Returns false if the file added, but could not attach to the parent.</returns>
        public bool AddFile(string FullFileName, string VaultFolderPath, bool UpdateExisting = true, bool DisableCheckInDesignFiles = true, string ParentFileToAttachTo = null)
        {
            //exclude CAD files with references
            System.IO.FileInfo mLocalFileInfo = new System.IO.FileInfo(FullFileName);
            if (DisableCheckInDesignFiles == true)
            {
                if (IsCadFile(mLocalFileInfo))
                {
                    return false;
                }
            }
            Autodesk.Connectivity.WebServicesTools.WebServiceManager mWsMgr = conn.WebServiceManager;

            ACW.Folder mFolder = mWsMgr.DocumentService.FindFoldersByPaths(new string[] { VaultFolderPath }).FirstOrDefault();
            if (mFolder.Id == -1)
            {
                return false;
            }
            string vaultFilePath = System.IO.Path.Combine(mFolder.FullName, mLocalFileInfo.Name).Replace("\\", "/");

            ACW.File wsFile = mWsMgr.DocumentService.FindLatestFilesByPaths(new string[] { vaultFilePath }).First();

            VDF.Currency.FilePathAbsolute vdfPath = new VDF.Currency.FilePathAbsolute(mLocalFileInfo.FullName);
            VDF.Vault.Currency.Entities.FileIteration vdfFile = null;
            //VDF.Vault.Currency.Entities.FileIteration addedFile = null;
            VDF.Vault.Currency.Entities.FileIteration mUploadedFile = null;
            if (wsFile.Id == -1)
            {
                // add new file to Vault
                var folderEntity = new Autodesk.DataManagement.Client.Framework.Vault.Currency.Entities.Folder(conn, mFolder);
                try
                {
                    mUploadedFile = conn.FileManager.AddFile(folderEntity, "Added by iLogic-Vault rule", null, null, ACW.FileClassification.None, false, vdfPath);
                }
                catch (Exception)
                {
                    return false;
                }

            }
            else
            {
                if (UpdateExisting == true)
                {
                    // checkin new file version
                    VDF.Vault.Settings.AcquireFilesSettings mCheckOutSettings = new VDF.Vault.Settings.AcquireFilesSettings(conn)
                    {
                        DefaultAcquisitionOption = VDF.Vault.Settings.AcquireFilesSettings.AcquisitionOption.Checkout
                    };
                    vdfFile = new VDF.Vault.Currency.Entities.FileIteration(conn, wsFile);
                    mCheckOutSettings.AddEntityToAcquire(vdfFile);
                    var results = conn.FileManager.AcquireFiles(mCheckOutSettings);
                    try
                    {
                        mUploadedFile = conn.FileManager.CheckinFile(results.FileResults.First().File, "Created by iLogic-Vault rule", false, null, null, false, null, ACW.FileClassification.None, false, vdfPath);
                    }
                    catch (Exception)
                    {
                        return false;
                    }
                }
            }

            if (ParentFileToAttachTo == null)
            {
                return true;
            }
            else
            {
                //check that the file can be found using the path and that it is available for checkout
                ACW.File mParentFile = mWsMgr.DocumentService.FindLatestFilesByPaths(new string[] { ParentFileToAttachTo }).First();
                if (mParentFile.Id == -1)
                {
                    return false;
                }
                else
                {
                    VDF.Vault.Currency.Entities.FileIteration mParFileIteration = new VDF.Vault.Currency.Entities.FileIteration(conn, mParentFile);
                    //try to check-out the file;
                    try
                    {
                        VDF.Vault.Settings.AcquireFilesSettings mCheckOutSettings = new VDF.Vault.Settings.AcquireFilesSettings(conn)
                        {
                            DefaultAcquisitionOption = VDF.Vault.Settings.AcquireFilesSettings.AcquisitionOption.Checkout
                        };
                        mCheckOutSettings.CheckoutComment = "iLogic-Vault Rule is about attaching an uploaded file.";
                        mCheckOutSettings.AddEntityToAcquire(mParFileIteration);
                        //Check-out and evaluate the results
                        VDF.Vault.Results.AcquireFilesResults acquireFilesResults = conn.FileManager.AcquireFiles(mCheckOutSettings);
                        if (acquireFilesResults.IsCancelled != true && acquireFilesResults.FileResults.First().Status == VDF.Vault.Results.FileAcquisitionResult.AcquisitionStatus.Success)
                        {
                            VDF.Vault.Currency.Entities.FileIteration mNewFileIteration = acquireFilesResults.FileResults.First().NewFileIteration;

                            //capture existing file associations
                            VDF.Vault.Settings.FileRelationshipGatheringSettings fileRelationshipGatheringSettings = new VDF.Vault.Settings.FileRelationshipGatheringSettings();
                            fileRelationshipGatheringSettings.IncludeAttachments = true;
                            fileRelationshipGatheringSettings.IncludeChildren = true;
                            fileRelationshipGatheringSettings.IncludeRelatedDocumentation = true;
                            fileRelationshipGatheringSettings.IncludeParents = false;
                            fileRelationshipGatheringSettings.IncludeHiddenEntities = true;
                            fileRelationshipGatheringSettings.IncludeLibraryContents = true;

                            IEnumerable<ACW.FileAssocLite> fileAssocsLite = conn.FileManager.GetFileAssociationLites(new long[] { mNewFileIteration.EntityIterationId },
                                fileRelationshipGatheringSettings);

                            //collect association parameters of existing file associations
                            List<ACW.FileAssocParam> fileAssocParams = new List<ACW.FileAssocParam>();
                            foreach (ACW.FileAssocLite item in fileAssocsLite)
                            {
                                //filter the parent associations and existing attachment of the uploaded file
                                if (item.ParFileId == mNewFileIteration.EntityIterationId && item.CldFileId != mUploadedFile.EntityIterationId)
                                {
                                    ACW.FileAssocParam param = new ACW.FileAssocParam();
                                    param.Typ = item.Typ;
                                    param.RefId = item.RefId;
                                    param.Source = item.Source;
                                    param.CldFileId = item.CldFileId;
                                    param.ExpectedVaultPath = item.ExpectedVaultPath;
                                    fileAssocParams.Add(param);
                                }
                            }

                            //create new association parameter of file to attach
                            ACW.FileAssocParam mNewParam = new ACW.FileAssocParam();
                            mNewParam.Typ = ACW.AssociationType.Attachment;
                            mNewParam.RefId = "";
                            mNewParam.Source = "";
                            mNewParam.CldFileId = mUploadedFile.EntityIterationId;
                            mNewParam.ExpectedVaultPath = mFolder.FullName + "/" + mUploadedFile.EntityName;

                            //combine the new parameter and the existing ones
                            fileAssocParams.Add(mNewParam);
                            ACW.FileAssocParam[] mFileAssocParamArray = (ACW.FileAssocParam[])fileAssocParams.ToArray();

                            //check-in the file providing the updated associations;
                            System.IO.Stream stream = null;
                            VDF.Vault.Currency.Entities.FileIteration mUpdatedParent;
                            try
                            {
                                //mUpdatedParent = conn.FileManager.CheckinFile(mNewFileIteration, "iLogic-Vault Rule attached " + mUploadedFile.EntityName, false, DateTime.Now, mFileAssocParamArray,
                                //                               null, true, mNewFileIteration.EntityName, mNewFileIteration.FileClassification, false, stream);
                                mUpdatedParent = conn.FileManager.CheckinFile(mNewFileIteration, "iLogic Rule modified Attachments", false, mFileAssocParamArray, null, true, mNewFileIteration.EntityName, mNewFileIteration.FileClassification, false);
                            }
                            catch (Exception)
                            {
                                mUpdatedParent = conn.FileManager.UndoCheckoutFile(mNewFileIteration, stream);
                                return false;
                            }
                        }
                    }
                    catch (Exception)
                    {
                        return false;
                    }
                }
                return true;
            }
        }


        /// <summary>
        /// Search multiple file(s) by 1 to many search criteria as property/value pairs. 
        /// Returns array of file names found, matching the criteria.
        /// Preset Search Operator Options: [Property] is (exactly) [Value]; multiple conditions link up using AND/OR condition, depending MatchAllCriteria = True/False
        /// </summary>
        /// <param name="SearchCriteria"></param>
        /// <param name="MatchAllCriteria"></param>
        /// <param name="FoldersSearched"></param>
        /// <returns></returns>
        public IList<string> CheckFilesExistBySearchCriteria(Dictionary<string, string> SearchCriteria, bool MatchAllCriteria = true, string[] FoldersSearched = null)
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
                foreach (ACW.File wsFile in totalResults)
                {
                    mFilesFound.Add(wsFile.Name);
                }
                return mFilesFound;
            }
            else
            {
                return null;
            }
        }


        /// <summary>
        /// Search multiple file(s) by 1 to many search criteria as property/value pairs. 
        /// Returns array of file names found, matching the criteria.
        /// Preset Search Operator Options: [Property] is (exactly) [Value]; multiple conditions link up using AND/OR condition, depending MatchAllCriteria = True/False
        /// </summary>
        /// <param name="SearchCriteria">Dictionary of property/value pairs</param>
        /// <param name="AllFilesVaultFileProperties">Dictonary of files' property dictionaries</param>
        /// <param name="MatchAllCriteria">Optional. Switches AND/OR conditions using multiple criterias. Default is true</param>
        /// <param name="FoldersSearched">Optional. Limit search scope to given folder path(s).</param>
        /// <returns>Array of file names found</returns>
        public IList<string> CheckFilesExistBySearchCriteria(Dictionary<string, string> SearchCriteria, ref Dictionary<string, Dictionary<string, string>> AllFilesVaultFileProperties,
            bool MatchAllCriteria = true, string[] FoldersSearched = null)
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
                foreach (ACW.File wsFile in totalResults)
                {
                    mFilesFound.Add(wsFile.Name);
                    Dictionary<string, string> mVaultFileProperties = new Dictionary<string, string>();
                    VDF.Vault.Currency.Entities.FileIteration mFileIt = new VDF.Vault.Currency.Entities.FileIteration(conn, wsFile);
                    mGetFileProps(mFileIt, ref mVaultFileProperties);
                    AllFilesVaultFileProperties.Add(mFileIt.EntityName, (mVaultFileProperties));
                }
                return mFilesFound;
            }
            else
            {
                return null;
            }
        }


        /// <summary>
        /// Get the local file's status in Vault.
        /// Validate the ErrorState = "None" to get all return values for vaulted files.
        /// Validate the ErrorState = (LocalFileNotFoundVaultFileNotFound|VaultFileNotFound) to validate files before first time check-in
        /// </summary>
        /// <param name="LocalFullFileName">Local path and file name, e.g., ThisDoc.FullFileName</param>
        /// <returns>ErrorState only if file is not added to Vault yet; otherwise Vault's default file status enumerations of CheckOutState, ConsumableState, ErrorState, LocalEditsState, LockState, RevisionState, VersionState</returns>
        public Dictionary<string, string> GetVaultFileStatus(string LocalFullFileName)
        {
            Dictionary<string, string> keyValues = new Dictionary<string, string>();

            //convert the local path to the corresponding Vault path; note - the file might be a virtual (to be created in future) one
            System.IO.FileInfo fileInfo = new System.IO.FileInfo(LocalFullFileName);
            string mVltFullFileName = null;
            string mWf = conn.WorkingFoldersManager.GetWorkingFolder("$/").FullPath;
            if (LocalFullFileName.ToLower().Contains(mWf.ToLower()))
                mVltFullFileName = LocalFullFileName.Replace(mWf, "$/");
            mVltFullFileName = mVltFullFileName.Replace("\\", "/");

            //get the file object consuming the Vault Path; if the file does not exist return the VaultFileNotFound status information; it is a custom ErrorState info
            Autodesk.Connectivity.WebServicesTools.WebServiceManager mWsMgr = conn.WebServiceManager;
            ACW.File mFile = mWsMgr.DocumentService.FindLatestFilesByPaths(new string[] { mVltFullFileName }).FirstOrDefault();

            if (mFile.Id == -1) // file not found locally and in Vault
            {
                if (!fileInfo.Exists)
                {
                    keyValues.Add("ErrorState", "LocalFileNotFoundVaultFileNotFound");
                }
                else
                {
                    keyValues.Add("ErrorState", "VaultFileNotFound");
                }
                return keyValues;
            }

            VDF.Vault.Currency.Entities.FileIteration mFileIteration = new VDF.Vault.Currency.Entities.FileIteration(conn, mFile);

            PropertyDefinitionDictionary mProps = conn.PropertyManager.GetPropertyDefinitions(VDF.Vault.Currency.Entities.EntityClassIds.Files, null, PropertyDefinitionFilter.IncludeAll);

            PropertyDefinition mVaultStatus = mProps[PropertyDefinitionIds.Client.VaultStatus];

            EntityStatusImageInfo status = conn.PropertyManager.GetPropertyValue(mFileIteration, mVaultStatus, null) as EntityStatusImageInfo;
            keyValues.Add("CheckOutState", status.Status.CheckoutState.ToString());
            keyValues.Add("ConsumableState", status.Status.ConsumableState.ToString());
            keyValues.Add("ErrorState", status.Status.ErrorState.ToString());
            keyValues.Add("LocalEditsState", status.Status.LocalEditsState.ToString());
            keyValues.Add("LockState", status.Status.LockState.ToString());
            keyValues.Add("RevisionState", status.Status.RevisionState.ToString());
            keyValues.Add("VersionState", status.Status.VersionState.ToString());

            return keyValues;
        }


        /// <summary>
        /// Convert the local file or folder path to the corresponding path in Vault.
        /// </summary>
        /// <param name="LocalPath">File or Folder path in local working folder</param>
        /// <returns>Vault Folder Path; if LocalPath is a file path, the file's parent Folder Path returns</returns>
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


        /// <summary>
        /// Downloads Vault file using full file path, e.g. "$/Designs/Base.ipt". Returns full file name in local working folder,
        /// download options include children and attachments; 
        /// </summary>
        /// <param name="VaultFullFileName">The full path and file name in Vault virtual folder structure, e.g., '$/Designs/Part1.ipt'</param>
        /// <param name="CheckOut">Optional. File downloaded does NOT check-out as default.</param>
        /// <returns>Local path/filename</returns>
        public string GetFileByFullFilePath(string VaultFullFileName, bool CheckOut = false)
        {
            List<string> mFiles = new List<string>();
            List<String> mFilesDownloaded = new List<string>();
            mFiles.Add(VaultFullFileName);
            ACW.File[] wsFiles = conn.WebServiceManager.DocumentService.FindLatestFilesByPaths(mFiles.ToArray());
            VDF.Vault.Currency.Entities.FileIteration mFileIt = new VDF.Vault.Currency.Entities.FileIteration(conn, (wsFiles[0]));

            if (mFileIt.EntityMasterId != -1)
            {
                //define download options, including DefaultAcquisitionOptions
                VDF.Vault.Settings.AcquireFilesSettings settings = CreateAcquireSettings(false);
                settings.AddFileToAcquire(mFileIt, settings.DefaultAcquisitionOption);

                //download
                VDF.Vault.Results.AcquireFilesResults results = conn.FileManager.AcquireFiles(settings);
                
                //capture primary file name for return (download may include children and attachments)
                if (results.FileResults != null)
                {
                    if (results.FileResults.Any(n => n.File.EntityName == mFileIt.EntityName))
                    {
                        mFilesDownloaded.Add(conn.WorkingFoldersManager.GetPathOfFileInWorkingFolder(mFileIt).FullPath.ToString());
                    }
                }
                //the download cancelled if the file already exists in the working folder
                if (results.IsCancelled == true)
                {
                    PropertyDefinitionDictionary mProps = conn.PropertyManager.GetPropertyDefinitions(VDF.Vault.Currency.Entities.EntityClassIds.Files, null, PropertyDefinitionFilter.IncludeAll);

                    PropertyDefinition mVaultStatus = mProps[PropertyDefinitionIds.Client.VaultStatus];

                    EntityStatusImageInfo mStatus = conn.PropertyManager.GetPropertyValue(mFileIt, mVaultStatus, null) as EntityStatusImageInfo;
                    if (mStatus.Status.ConsumableState == EntityStatus.ConsumableStateEnum.LatestConsumable)
                    {
                        mFilesDownloaded.Add(conn.WorkingFoldersManager.GetPathOfFileInWorkingFolder(mFileIt).FullPath.ToString());
                    }
                }

                //define checkout options and checkout
                if (CheckOut)
                {
                    settings = CreateAcquireSettings(true);
                    settings.AddFileToAcquire(mFileIt, settings.DefaultAcquisitionOption);
                    results = conn.FileManager.AcquireFiles(settings);
                }

                //return the file
                if (mFilesDownloaded.Count > 0)
                {
                    return mFilesDownloaded[0];
                }
                else
                {
                    return null;
                }
            }
            return null;
        }

        /// <summary>
        /// Downloads Vault file using full file path, e.g. "$/Designs/Base.ipt". Returns full file name in local working folder,
        /// download options include children and attachments; 
        /// File Properties return all values converted to text. Access the value using the Vault property display name as key.
        /// </summary>
        /// <param name="VaultFullFileName">The full path and file name in Vault virtual folder structure, e.g., '$/Designs/Part1.ipt'</param>
        /// <param name="VaultFileProperties">pairs of Vault File property display name and property value</param>
        /// <param name="CheckOut">Default value is 'False'; set to 'True' to check-out the downloaded file</param>
        /// <returns>Local path/filename</returns>
        public string GetFileByFullFilePath(string VaultFullFileName, ref Dictionary<string, string> VaultFileProperties, bool CheckOut = false)
        {
            List<string> mFiles = new List<string>();
            List<String> mFilesDownloaded = new List<string>();
            mFiles.Add(VaultFullFileName);
            ACW.File[] wsFiles = conn.WebServiceManager.DocumentService.FindLatestFilesByPaths(mFiles.ToArray());
            VDF.Vault.Currency.Entities.FileIteration mFileIt = new VDF.Vault.Currency.Entities.FileIteration(conn, (wsFiles[0]));

            if (mFileIt.EntityMasterId != -1)
            {
                //define download options, including DefaultAcquisitionOptions
                VDF.Vault.Settings.AcquireFilesSettings settings = CreateAcquireSettings(false);
                settings.AddFileToAcquire(mFileIt, settings.DefaultAcquisitionOption);

                //download
                VDF.Vault.Results.AcquireFilesResults results = conn.FileManager.AcquireFiles(settings);

                //capture primary file name for return (download may include children and attachments)
                if (results.FileResults != null)
                {
                    if (results.FileResults.Any(n => n.File.EntityName == mFileIt.EntityName))
                    {
                        mFilesDownloaded.Add(conn.WorkingFoldersManager.GetPathOfFileInWorkingFolder(mFileIt).FullPath.ToString());
                    }
                }
                //the download cancelled if the file already exists in the working folder
                if (results.IsCancelled == true)
                {
                    PropertyDefinitionDictionary mProps = conn.PropertyManager.GetPropertyDefinitions(VDF.Vault.Currency.Entities.EntityClassIds.Files, null, PropertyDefinitionFilter.IncludeAll);

                    PropertyDefinition mVaultStatus = mProps[PropertyDefinitionIds.Client.VaultStatus];

                    EntityStatusImageInfo mStatus = conn.PropertyManager.GetPropertyValue(mFileIt, mVaultStatus, null) as EntityStatusImageInfo;
                    if (mStatus.Status.ConsumableState == EntityStatus.ConsumableStateEnum.LatestConsumable)
                    {
                        mFilesDownloaded.Add(conn.WorkingFoldersManager.GetPathOfFileInWorkingFolder(mFileIt).FullPath.ToString());
                    }
                }

                //define checkout options and checkout
                if (CheckOut)
                {
                    settings = CreateAcquireSettings(true);
                    settings.AddFileToAcquire(mFileIt, settings.DefaultAcquisitionOption);
                    results = conn.FileManager.AcquireFiles(settings);
                }

                //return the file and optional output
                if (mFilesDownloaded.Count > 0)
                {
                    //collect file properties
                    mGetFileProps(mFileIt, ref VaultFileProperties);

                    return mFilesDownloaded[0];
                }
                else
                {
                    return null;
                }
            }

            return null;
        }

        /// <summary>
        /// Downloads Vault file using full file path, e.g. "$/Designs/Base.ipt". Returns full file name in local working folder,
        /// download options include children and attachments; 
        /// File and Item property dictionaries return all values converted to text. Access the value using the Vault property display name as key.
        /// </summary>
        /// <param name="VaultFullFileName">The full path and file name in Vault virtual folder structure, e.g., '$/Designs/Part1.ipt'</param>
        /// <param name="VaultFileProperties">pairs of Vault File property display name and property value</param>
        /// <param name="VaultItemProperties">pairs of Vault Item property display name and property value</param>
        /// <param name="CheckOut"></param>
        /// <returns>Local path/filename</returns>
        public string GetFileByFullFilePath(string VaultFullFileName, ref Dictionary<string, string> VaultFileProperties,
            ref Dictionary<string, string> VaultItemProperties, bool CheckOut = false)
        {
            if (IsVaultBasic)
            {
                MessageBox.Show("The iLogic-Vault method GetFileByFullFilePath overload including Item properties is not available for Vault Basic", "iLogic-Vault", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return null;
            }

            List<string> mFiles = new List<string>();
            List<String> mFilesDownloaded = new List<string>();
            mFiles.Add(VaultFullFileName);
            ACW.File[] wsFiles = conn.WebServiceManager.DocumentService.FindLatestFilesByPaths(mFiles.ToArray());
            VDF.Vault.Currency.Entities.FileIteration mFileIt = new VDF.Vault.Currency.Entities.FileIteration(conn, (wsFiles[0]));

            if (mFileIt.EntityMasterId != -1)
            {
                //define download options, including DefaultAcquisitionOptions
                VDF.Vault.Settings.AcquireFilesSettings settings = CreateAcquireSettings(false);
                settings.AddFileToAcquire(mFileIt, settings.DefaultAcquisitionOption);

                //download
                VDF.Vault.Results.AcquireFilesResults results = conn.FileManager.AcquireFiles(settings);

                //capture primary file name for return (download may include children and attachments)
                if (results.FileResults != null)
                {
                    if (results.FileResults.Any(n => n.File.EntityName == mFileIt.EntityName))
                    {
                        mFilesDownloaded.Add(conn.WorkingFoldersManager.GetPathOfFileInWorkingFolder(mFileIt).FullPath.ToString());
                    }
                }
                //the download cancelled if the file already exists in the working folder
                if (results.IsCancelled == true)
                {
                    PropertyDefinitionDictionary mProps = conn.PropertyManager.GetPropertyDefinitions(VDF.Vault.Currency.Entities.EntityClassIds.Files, null, PropertyDefinitionFilter.IncludeAll);

                    PropertyDefinition mVaultStatus = mProps[PropertyDefinitionIds.Client.VaultStatus];

                    EntityStatusImageInfo mStatus = conn.PropertyManager.GetPropertyValue(mFileIt, mVaultStatus, null) as EntityStatusImageInfo;
                    if (mStatus.Status.ConsumableState == EntityStatus.ConsumableStateEnum.LatestConsumable)
                    {
                        mFilesDownloaded.Add(conn.WorkingFoldersManager.GetPathOfFileInWorkingFolder(mFileIt).FullPath.ToString());
                    }
                }

                //define checkout options and checkout
                if (CheckOut)
                {
                    settings = CreateAcquireSettings(true);
                    settings.AddFileToAcquire(mFileIt, settings.DefaultAcquisitionOption);
                    results = conn.FileManager.AcquireFiles(settings);
                }

                //return the file and optional output
                if (mFilesDownloaded.Count > 0)
                {
                    //collect file properties
                    mGetFileProps(mFileIt, ref VaultFileProperties);

                    //collect item properties if file linked to item
                    ACW.Item[] items = conn.WebServiceManager.ItemService.GetItemsByFileId(mFileIt.EntityIterationId);
                    if (items.Length > 0)
                    {
                        //todo: handle 1:n file item links (may happen with model states)
                        ACW.Item item = items[0];
                        mGetItemProps(item, ref VaultItemProperties);
                    }

                    return mFilesDownloaded[0];
                }
                else
                {
                    return null;
                }
            }

            return null;
        }


        /// <summary>
        /// Copy Vault file on file server and download using full file path, e.g. "$/Designs/Base.ipt".
        /// Create new file name using default or named numbering scheme.
        /// Preset Options: Download Children (recursively) = Enabled, Enforce Overwrite = True
        /// Not available for Vault Basic.
        /// </summary>
        /// <param name="VaultFullFileName">The full path and file name in Vault virtual folder structure, e.g., '$/Designs/Part1.ipt'</param>
        /// <param name="NumberingScheme">Individual scheme name or 'Default'</param>
        /// <param name="InputParams">Optional according scheme definition. User input values in order of scheme configuration</param>
        /// <param name="UpdatePartNumber">Optional. Update Part Number property to match new file name</param>
        /// <param name="CheckOut">Optional. File copy will check-out as default.</param>
        /// <returns>Local path/filename</returns>
        public string GetFileCopyBySourceFileNameAndAutoNumber(string VaultFullFileName, string NumberingScheme, string[] InputParams = null, bool CheckOut = true, bool UpdatePartNumber = true)
        {
            if (IsVaultBasic)
            {
                MessageBox.Show("The iLogic-Vault method GetFileCopyBySourceFileNameAndAutoNumber is not available for Vault Basic", "iLogic-Vault", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return null;
            }

            //Get Vault File object
            List<string> mFiles = new List<string>();
            List<String> mFilesDownloaded = new List<string>();
            mFiles.Add(VaultFullFileName);
            ACW.File mSourceFile = conn.WebServiceManager.DocumentService.FindLatestFilesByPaths(mFiles.ToArray()).First();

            if (mSourceFile.Name == null) return null;

            //build new file name
            string mNewNumber = GetNewNumber(NumberingScheme, InputParams);
            if (mNewNumber == null) return null;
            string mNewFileName = String.Format("{0}{1}{2}", mNewNumber, ".", (mSourceFile.Name).Split('.').Last());

            //create file iteration as copy from source
            VDF.Vault.Currency.Entities.FileIteration mFileIt = CreateFileCopy(mSourceFile, mNewFileName);

            //Optionally update Partnumber property
            if (UpdatePartNumber)
            {
                Dictionary<ACW.PropDef, object> mPropDictonary = new Dictionary<ACW.PropDef, object>();

                if (propDefs == null)
                {
                    propDefs = conn.WebServiceManager.PropertyService.GetPropertyDefinitionsByEntityClassId("FILE");
                }
                ACW.PropDef propDef = propDefs.SingleOrDefault(n => n.SysName == "PartNumber");
                mPropDictonary.Add(propDef, mNewNumber);

                mUpdateFileProperties((ACW.File)mFileIt, mPropDictonary);
                mFileIt = new VDF.Vault.Currency.Entities.FileIteration(conn, conn.WebServiceManager.DocumentService.GetLatestFileByMasterId(mFileIt.EntityMasterId));
            }

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
            return null;
        }

        /// <summary>
        /// Copy Vault file on file server and download using full file path, e.g. "$/Designs/Base.ipt".
        /// Create new file name re-using source file's extension and new file name variable.
        /// Preset Options: Download Children (recursively) = Enabled, Enforce Overwrite = True
        /// Not available for Vault Basic.
        /// </summary>
        /// <param name="VaultFullFileName">The full path and file name in Vault virtual folder structure, e.g., '$/Designs/Part1.ipt'</param>
        /// <param name="NewFileNameNoExt">New name without extension</param>
        /// <param name="CheckOut">Optional. File copy will check-out as default.</param>
        /// <param name="UpdatePartNumber">Optional. Update Part Number property to match new file name</param>
        /// <returns>Local path/filename</returns>
        public string GetFileCopyBySourceFileNameAndNewName(string VaultFullFileName, string NewFileNameNoExt, bool CheckOut = true, bool UpdatePartNumber = true)
        {
            if (IsVaultBasic)
            {
                MessageBox.Show("The iLogic-Vault method GetFileCopyBySourceFileNameAndNewName is not available for Vault Basic", "iLogic-Vault", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return null;
            }

            //get Vault File object
            List<string> mFiles = new List<string>();
            List<String> mFilesDownloaded = new List<string>();
            mFiles.Add(VaultFullFileName);
            ACW.File mSourceFile = conn.WebServiceManager.DocumentService.FindLatestFilesByPaths(mFiles.ToArray()).First();

            if (mSourceFile == null) return null;

            string mNewFileName = String.Format("{0}{1}{2}", NewFileNameNoExt, ".", (mSourceFile.Name).Split('.').Last());

            //create file iteration as copy from source
            VDF.Vault.Currency.Entities.FileIteration mFileIt = CreateFileCopy(mSourceFile, mNewFileName);

            //Optionally update Partnumber property
            if (UpdatePartNumber)
            {
                Dictionary<ACW.PropDef, object> mPropDictonary = new Dictionary<ACW.PropDef, object>();

                ACW.PropDef[] propDefs = conn.WebServiceManager.PropertyService.GetPropertyDefinitionsByEntityClassId("FILE");
                ACW.PropDef propDef = propDefs.SingleOrDefault(n => n.SysName == "PartNumber");
                mPropDictonary.Add(propDef, mNewFileName);

                mUpdateFileProperties(mFileIt, mPropDictonary);
                mFileIt = new VDF.Vault.Currency.Entities.FileIteration(conn, conn.WebServiceManager.DocumentService.GetLatestFileByMasterId(mFileIt.EntityMasterId));
            }

            //build the download settings
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
        /// Search for a file by 1 to many search criteria as property/value pairs. 
        /// Downloads the first file found, if the search result lists more than a single file. Dependents and attachments are included. Overwrites existing files.
        /// Preset Search Operator Options: [Property] is (exactly) [Value]; multiple conditions link up using AND/OR condition, depending MatchAllCriteria = True/False
        /// Returns the file name downloaded (does not return names of downloaded children and attachments). 
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

                //capture primary file name for return (download may include children and attachments)
                if (results.FileResults != null)
                {
                    if (results.FileResults.Any(n => n.File.EntityName == mFileIt.EntityName))
                    {
                        mFilesDownloaded.Add(conn.WorkingFoldersManager.GetPathOfFileInWorkingFolder(mFileIt).FullPath.ToString());
                    }
                }
                //the download cancelled if the file already exists in the working folder
                if (results.IsCancelled == true)
                {
                    PropertyDefinitionDictionary mProps = conn.PropertyManager.GetPropertyDefinitions(VDF.Vault.Currency.Entities.EntityClassIds.Files, null, PropertyDefinitionFilter.IncludeAll);

                    PropertyDefinition mVaultStatus = mProps[PropertyDefinitionIds.Client.VaultStatus];

                    EntityStatusImageInfo mStatus = conn.PropertyManager.GetPropertyValue(mFileIt, mVaultStatus, null) as EntityStatusImageInfo;
                    if (mStatus.Status.ConsumableState == EntityStatus.ConsumableStateEnum.LatestConsumable)
                    {
                        mFilesDownloaded.Add(conn.WorkingFoldersManager.GetPathOfFileInWorkingFolder(mFileIt).FullPath.ToString());
                    }
                }

                //define checkout options and checkout
                if (CheckOut)
                {
                    settings = CreateAcquireSettings(true);
                    settings.AddFileToAcquire(mFileIt, settings.DefaultAcquisitionOption);
                    results = conn.FileManager.AcquireFiles(settings);
                }

                //return the file and optional output
                if (mFilesDownloaded.Count > 0)
                {
                    return mFilesDownloaded[0];
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
        /// Search for a file by 1 to many search criteria as property/value pairs. 
        /// Downloads the first file found, if the search result lists more than a single file. Dependents and attachments are included. Overwrites existing files.
        /// Preset Search Operator Options: [Property] is (exactly) [Value]; multiple conditions link up using AND/OR condition, depending MatchAllCriteria = True/False
        /// Returns the file name downloaded (does not return names of downloaded children and attachments).
        /// File property dictionaries return all values converted to text. Access the value using the Vault property display name as key.
        /// </summary>
        /// <param name="SearchCriteria">Dictionary of property/value pairs</param>
        /// <param name="VaultFileProperties">pairs of Vault File property display name and property value</param>
        /// <param name="MatchAllCriteria">Optional. Switches AND/OR conditions using multiple criterias. Default is true</param>
        /// <param name="CheckOut">Optional. File downloaded does NOT check-out as default</param>
        /// <param name="FoldersSearched">Optional. Limit search scope to given folder path(s).</param>
        /// <returns>Local path/filename</returns>
        public string GetFileBySearchCriteria(Dictionary<string, string> SearchCriteria, ref Dictionary<string, string> VaultFileProperties, bool MatchAllCriteria = true, bool CheckOut = false, string[] FoldersSearched = null)
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

                //capture primary file name for return (download may include children and attachments)
                if (results.FileResults != null)
                {
                    if (results.FileResults.Any(n => n.File.EntityName == mFileIt.EntityName))
                    {
                        mFilesDownloaded.Add(conn.WorkingFoldersManager.GetPathOfFileInWorkingFolder(mFileIt).FullPath.ToString());
                    }
                }
                //the download cancelled if the file already exists in the working folder
                if (results.IsCancelled == true)
                {
                    PropertyDefinitionDictionary mProps = conn.PropertyManager.GetPropertyDefinitions(VDF.Vault.Currency.Entities.EntityClassIds.Files, null, PropertyDefinitionFilter.IncludeAll);

                    PropertyDefinition mVaultStatus = mProps[PropertyDefinitionIds.Client.VaultStatus];

                    EntityStatusImageInfo mStatus = conn.PropertyManager.GetPropertyValue(mFileIt, mVaultStatus, null) as EntityStatusImageInfo;
                    if (mStatus.Status.ConsumableState == EntityStatus.ConsumableStateEnum.LatestConsumable)
                    {
                        mFilesDownloaded.Add(conn.WorkingFoldersManager.GetPathOfFileInWorkingFolder(mFileIt).FullPath.ToString());
                    }
                }

                //define checkout options and checkout
                if (CheckOut)
                {
                    settings = CreateAcquireSettings(true);
                    settings.AddFileToAcquire(mFileIt, settings.DefaultAcquisitionOption);
                    results = conn.FileManager.AcquireFiles(settings);
                }

                //return the file and optional output
                if (mFilesDownloaded.Count > 0)
                {
                    //get file properties
                    mGetFileProps(mFileIt, ref VaultFileProperties);

                    return mFilesDownloaded[0];
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
        /// Search for a file by 1 to many search criteria as property/value pairs. 
        /// Downloads the first file found, if the search result lists more than a single file. Dependents and attachments are included. Overwrites existing files.
        /// Preset Search Operator Options: [Property] is (exactly) [Value]; multiple conditions link up using AND/OR condition, depending MatchAllCriteria = True/False
        /// Returns the file name downloaded (does not return names of downloaded children and attachments).
        /// File and Item property dictionaries return all values converted to text. Access the value using the Vault property display name as key.
        /// </summary>
        /// <param name="SearchCriteria">Dictionary of property/value pairs</param>
        /// <param name="VaultFileProperties">pairs of Vault File property display name and property value</param>
        /// <param name="VaultItemProperties">pairs of Vault Item property display name and property value</param>
        /// <param name="MatchAllCriteria">Optional. Switches AND/OR conditions using multiple criterias. Default is true</param>
        /// <param name="CheckOut">Optional. File downloaded does NOT check-out as default</param>
        /// <param name="FoldersSearched">Optional. Limit search scope to given folder path(s).</param>
        /// <returns>Local path/filename</returns>
        public string GetFileBySearchCriteria(Dictionary<string, string> SearchCriteria, ref Dictionary<string, string> VaultFileProperties,
            ref Dictionary<string, string> VaultItemProperties, bool MatchAllCriteria = true, bool CheckOut = false, string[] FoldersSearched = null)
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

                //capture primary file name for return (download may include children and attachments)
                if (results.FileResults != null)
                {
                    if (results.FileResults.Any(n => n.File.EntityName == mFileIt.EntityName))
                    {
                        mFilesDownloaded.Add(conn.WorkingFoldersManager.GetPathOfFileInWorkingFolder(mFileIt).FullPath.ToString());
                    }
                }
                //the download cancelled if the file already exists in the working folder
                if (results.IsCancelled == true)
                {
                    PropertyDefinitionDictionary mProps = conn.PropertyManager.GetPropertyDefinitions(VDF.Vault.Currency.Entities.EntityClassIds.Files, null, PropertyDefinitionFilter.IncludeAll);

                    PropertyDefinition mVaultStatus = mProps[PropertyDefinitionIds.Client.VaultStatus];

                    EntityStatusImageInfo mStatus = conn.PropertyManager.GetPropertyValue(mFileIt, mVaultStatus, null) as EntityStatusImageInfo;
                    if (mStatus.Status.ConsumableState == EntityStatus.ConsumableStateEnum.LatestConsumable)
                    {
                        mFilesDownloaded.Add(conn.WorkingFoldersManager.GetPathOfFileInWorkingFolder(mFileIt).FullPath.ToString());
                    }
                }

                //define checkout options and checkout
                if (CheckOut)
                {
                    settings = CreateAcquireSettings(true);
                    settings.AddFileToAcquire(mFileIt, settings.DefaultAcquisitionOption);
                    results = conn.FileManager.AcquireFiles(settings);
                }

                //return the file and optional output
                if (mFilesDownloaded.Count > 0)
                {
                    //get file properties
                    mGetFileProps(mFileIt, ref VaultFileProperties);

                    //collect item properties if file linked to item
                    ACW.Item[] items = conn.WebServiceManager.ItemService.GetItemsByFileId(mFileIt.EntityIterationId);
                    if (items.Length > 0)
                    {
                        //todo: handle 1:n file item links (may happen with model states)
                        ACW.Item item = items[0];
                        mGetItemProps(item, ref VaultItemProperties);
                    }

                    return mFilesDownloaded[0];
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
        /// Search for multiple files by 1 to many search criteria as property/value pairs. 
        /// Downloads all files found, matching the criteria. Dependents and attachments are included. Overwrites existing files.
        /// Preset Search Operator Options: [Property] is (exactly) [Value]; multiple conditions link up using AND/OR condition, depending MatchAllCriteria = True/False
        /// Returns list of files names downloaded (does not return names of downloaded children and attachments).
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

                //capture primary file names for return (we don't return names of downloaded children and attachment files)
                if (results.FileResults != null)
                {
                    foreach (VDF.Vault.Currency.Entities.FileIteration mFileIt in mFilesFound)
                    {
                        if (results.FileResults.Any(n => n.File.EntityName == mFileIt.EntityName))
                        {
                            mFilesDownloaded.Add(conn.WorkingFoldersManager.GetPathOfFileInWorkingFolder(mFileIt).FullPath.ToString());
                        }
                    }
                }
                //the download cancelled if the file already exists in the working folder
                if (results.IsCancelled == true)
                {
                    foreach (VDF.Vault.Currency.Entities.FileIteration mFileIt in mFilesFound)
                    {
                        PropertyDefinitionDictionary mProps = conn.PropertyManager.GetPropertyDefinitions(VDF.Vault.Currency.Entities.EntityClassIds.Files, null, PropertyDefinitionFilter.IncludeAll);

                        PropertyDefinition mVaultStatus = mProps[PropertyDefinitionIds.Client.VaultStatus];

                        EntityStatusImageInfo mStatus = conn.PropertyManager.GetPropertyValue(mFileIt, mVaultStatus, null) as EntityStatusImageInfo;
                        if (mStatus.Status.ConsumableState == EntityStatus.ConsumableStateEnum.LatestConsumable)
                        {
                            mFilesDownloaded.Add(conn.WorkingFoldersManager.GetPathOfFileInWorkingFolder(mFileIt).FullPath.ToString());
                        }
                    }
                }

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

                //return the files
                if (mFilesDownloaded.Count > 0)
                {
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
        /// Search multiple file(s) by 1 to many search criteria as property/value pairs. 
        /// Downloads first file matching all or any search criterias.
        /// Preset Search Operator Options: [Property] is (exactly) [Value]; multiple conditions link up using AND condition.
        /// Preset Download Options: Download Children (recursively) = Enabled, Enforce Overwrite = True
        /// Not available for Vault Basic.
        /// </summary>
        /// <param name="SearchCriteria">Dictionary of property/value pairs</param>        
        /// <param name="NumberingScheme">Individual scheme name or 'Default'</param>
        /// <param name="MatchAllCriteria">Optional. Switches AND/OR conditions using multiple criterias. Default is true</param>
        /// <param name="InputParams">Optional according scheme definition. User input values in order of scheme configuration</param>
        /// <param name="CheckOut">Optional. File copy will check-out as default.</param>
        /// <param name="FoldersSearched">Optional. Limit search scope to given folder path(s).</param>
        /// <param name="UpdatePartNumber">Optional. Update Part Number property to match new file name</param>
        /// <returns>Local path/filenamen</returns>
        public string GetFileCopyBySourceFileSearchAndAutoNumber(Dictionary<string, string> SearchCriteria, string NumberingScheme, bool MatchAllCriteria = true, string[] InputParams = null, bool CheckOut = true, string[] FoldersSearched = null, bool UpdatePartNumber = true)
        {
            if (IsVaultBasic)
            {
                MessageBox.Show("The iLogic-Vault method GetFileCopyBySourceFileSearchAndAutoNumber is not available for Vault Basic", "iLogic-Vault", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return null;
            }

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
                ACW.File mSourceFile = totalResults.First<ACW.File>();
                if (mSourceFile == null) return null;

                //get new file name/number
                string mNewNumber = GetNewNumber(NumberingScheme, InputParams);
                if (mNewNumber == null) return null;
                string mNewFileName = String.Format("{0}{1}{2}", mNewNumber, ".", (mSourceFile.Name).Split('.').Last());

                //create new file iteration as copy of source file
                VDF.Vault.Currency.Entities.FileIteration mFileIt = CreateFileCopy(mSourceFile, mNewFileName);

                //Optionally pdate Partnumber property
                if (UpdatePartNumber)
                {
                    Dictionary<ACW.PropDef, object> mPropDictonary = new Dictionary<ACW.PropDef, object>();

                    ACW.PropDef[] propDefs = conn.WebServiceManager.PropertyService.GetPropertyDefinitionsByEntityClassId("FILE");
                    ACW.PropDef propDef = propDefs.SingleOrDefault(n => n.SysName == "PartNumber");
                    mPropDictonary.Add(propDef, mNewNumber);

                    mUpdateFileProperties((ACW.File)mFileIt, mPropDictonary);
                    mFileIt = new VDF.Vault.Currency.Entities.FileIteration(conn, conn.WebServiceManager.DocumentService.GetLatestFileByMasterId(mFileIt.EntityMasterId));
                }

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
        /// Search multiple file(s) by 1 to many search criteria as property/value pairs. 
        /// Downloads first file matching all or any search criterias.
        /// Preset Search Operator Options: [Property] is (exactly) [Value]; multiple conditions link up using AND condition.
        /// Preset Download Options: Download Children (recursively) = Enabled, Enforce Overwrite = True
        /// Not available for Vault Basic.
        /// </summary>
        /// <param name="SearchCriteria">Dictionary of property/value pairs</param>        
        /// <param name="NewFileNameNoExt">New name without extension</param>
        /// <param name="MatchAllCriteria">Optional. Switches AND/OR conditions using multiple criterias. Default is true</param>
        /// <param name="CheckOut">Optional. File copy will check-out as default.</param>
        /// <param name="FoldersSearched">Optional. Limit search scope to given folder path(s).</param>
        /// <param name="UpdatePartNumber">Optional. Update Part Number property to match new file name</param>
        /// <returns>Local path/filename</returns>
        public string GetFileCopyBySourceFileSearchAndNewName(Dictionary<string, string> SearchCriteria, string NewFileNameNoExt, bool MatchAllCriteria = true, bool CheckOut = true, string[] FoldersSearched = null, bool UpdatePartNumber = true)
        {
            if (IsVaultBasic)
            {
                MessageBox.Show("The iLogic-Vault method GetFileCopyBySourceFileSearchAndNewName is not available for Vault Basic", "iLogic-Vault", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return null;
            }

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
                ACW.File mSourceFile = totalResults.First<ACW.File>();
                if (mSourceFile == null) return null;

                string mNewFileName = String.Format("{0}{1}{2}", NewFileNameNoExt, ".", (mSourceFile.Name).Split('.').Last());
                VDF.Vault.Currency.Entities.FileIteration mFileIt = CreateFileCopy(mSourceFile, mNewFileName);

                //Optionally pdate Partnumber property
                if (UpdatePartNumber)
                {
                    Dictionary<ACW.PropDef, object> mPropDictonary = new Dictionary<ACW.PropDef, object>();

                    ACW.PropDef[] propDefs = conn.WebServiceManager.PropertyService.GetPropertyDefinitionsByEntityClassId("FILE");
                    ACW.PropDef propDef = propDefs.SingleOrDefault(n => n.SysName == "PartNumber");
                    mPropDictonary.Add(propDef, mNewFileName);

                    mUpdateFileProperties(mFileIt, mPropDictonary);
                    mFileIt = new VDF.Vault.Currency.Entities.FileIteration(conn, conn.WebServiceManager.DocumentService.GetLatestFileByMasterId(mFileIt.EntityMasterId));
                }

                //build download options, including DefaultAcquisitionOptions
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
        /// Create single file number by scheme name and optional input parameters
        /// </summary>
        /// <param name="mSchmName">Name of individual Numbering Scheme or "Default" for pre-set scheme</param>
        /// <param name="mSchmPrms">User input parameter in order of scheme configuration</param>
        /// <returns>new number</returns>
        public string GetNewNumber(string mSchmName, string[] mSchmPrms = null)
        {
            if (IsVaultBasic)
            {
                MessageBox.Show("The iLogic-Vault method GetNewNumber is not available for Vault Basic", "iLogic-Vault", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return null;
            }

            ACW.NumSchm NmngSchm = null;
            try
            {
                if (mSchmName == "Default")
                {
                    NmngSchm = conn.WebServiceManager.NumberingService.GetNumberingSchemes("FILE", ACW.NumSchmType.Activated).First(n => n.IsDflt == true);
                }
                else
                {
                    NmngSchm = conn.WebServiceManager.NumberingService.GetNumberingSchemes("FILE", ACW.NumSchmType.Activated).First(n => n.Name == mSchmName);
                }
                return conn.WebServiceManager.DocumentService.GenerateFileNumber(NmngSchm.SchmID, mSchmPrms);
            }
            catch (Exception)
            {
                return null;
            }
        }

        /// <summary>
        /// Copies a local file to a new name. 
        /// The source file's location and extension are captured and apply to the copy.
        /// Use Check-In (iLogic) command for adding the new file to Vault.
        /// </summary>
        /// <param name="mFullFileName">File name including full path</param>
        /// <param name="mNewNameNoExtension">The new target name of the copied file. Path and extension will transfer from the source file.</param>
        /// <returns>Local path/filename</returns>
        public string CopyLocalFile(string mFullFileName, string mNewNameNoExtension)
        {
            try
            {
                System.IO.FileInfo mFileInfo = new System.IO.FileInfo(mFullFileName);
                string mExt = mFileInfo.Extension;
                string mFileName = mFileInfo.Name;
                mFileName = mFileName.Replace(mExt, "");
                string mNewFullFileName = mFullFileName.Replace(mFileName, mNewNameNoExtension);
                System.IO.File.Copy(mFullFileName, mNewFullFileName, true);
                System.IO.FileInfo mNewFileInfo = new System.IO.FileInfo(mNewFullFileName);
                return mNewFileInfo.FullName;
            }
            catch (Exception)
            {
                return null;
            }
        }


        /// <summary>
        /// Download Thumbnail Image of the given file as Image file.
        /// </summary>
        /// <param name="VaultFullFileName">The full path and file name in Vault virtual folder structure, e.g., '$/Designs/Part1.ipt'</param>
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

        /// <summary>
        /// Get Thumbnail of the given file as Image object.
        /// </summary>
        /// <param name="VaultFullFileName">The full path and file name in Vault virtual folder structure, e.g., '$/Designs/Part1.ipt'</param>
        /// <param name="Width">Optional. Image pixel size</param>
        /// <param name="Height">Optional. Image pixel size.</param>
        /// <returns>System.Drawing.Image object</returns>
        public Image GetThumbnailImageByFullSourceFilePath(string VaultFullFileName, int Width = 300, int Height = 300)
        {
            List<string> mFiles = new List<string>();
            mFiles.Add(VaultFullFileName);
            try
            {
                ACW.File[] wsFiles = conn.WebServiceManager.DocumentService.FindLatestFilesByPaths(mFiles.ToArray());
                VDF.Vault.Currency.Entities.FileIteration mFileIt = new VDF.Vault.Currency.Entities.FileIteration(conn, (wsFiles[0]));

                Image image = GetThumbnailImage(mFileIt, Width, Height);
                if (image != null)
                {
                    return image;
                }
                else
                {
                    return null;
                }
            }
            catch (Exception)
            {
                return null;
            }
        }

        /// <summary>
        /// Get Thumbnail Image of the file searched as Image object.
        /// </summary>
        /// <param name="SearchCriteria">Dictionary of property/value pairs</param>
        /// <param name="MatchAllCriteria">Optional. Switches AND/OR conditions using multiple criterias. Default is true</param>
        /// <param name="FoldersSearched">Optional. Limit search scope to given folder path(s).</param>
        /// <param name="Width">Optional. Image pixel size</param>
        /// <param name="Height">Optional. Image pixel size.</param>
        /// <returns>System.Drawing.Image object</returns>
        public Image GetThumbnailImageBySearchCriteria(Dictionary<string, string> SearchCriteria, bool MatchAllCriteria = true, string[] FoldersSearched = null, int Width = 300, int Height = 300)
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
                    ACW.File wsFile = totalResults.First<ACW.File>();
                    VDF.Vault.Currency.Entities.FileIteration mFileIt = new VDF.Vault.Currency.Entities.FileIteration(conn, (wsFile));

                    Image image = GetThumbnailImage(mFileIt, Width, Height);
                    if (image != null)
                    {
                        return image;
                    }
                    else
                    {
                        return null;
                    }
                }
                catch (Exception)
                {
                    return null;
                }
            }
            return null;
        }


        /// <summary>
        /// Update multiple Vault file properties. Not available for Vault Basic.
        /// </summary>
        /// <param name="VaultFullFileName">The full path and file name in Vault virtual folder structure, e.g., '$/Designs/Part1.ipt'</param>
        /// <param name="VaultFileProperties">pairs of Vault File property display name and property value of respective type</param>
        /// <param name="GetLatestFile">Option to download the updated file from Vault; default is true.</param>
        /// <returns>Returns true on success; returns false in case of failure, e.g., if the file is not available for check out</returns>
        public bool UpdateVaultFileProperties(string VaultFullFileName, Dictionary<string, object> VaultFileProperties, bool GetLatestFile = true)
        {
            //this method is available for Vault Workgroup and Professional only.
            if (IsVaultBasic)
            {
                MessageBox.Show("The iLogic-Vault method UpdateVaultFileProperties is not available for Vault Basic", "iLogic-Vault", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }
            //convert the given file path to the file object
            List<string> mFiles = new List<string>();
            List<String> mFilesDownloaded = new List<string>();
            mFiles.Add(VaultFullFileName);
            ACW.File[] wsFiles = conn.WebServiceManager.DocumentService.FindLatestFilesByPaths(mFiles.ToArray());
            ACW.File mFile = wsFiles.FirstOrDefault();
            VDF.Vault.Currency.Entities.FileIteration mFileIt = new VDF.Vault.Currency.Entities.FileIteration(conn, (wsFiles[0]));

            Dictionary<ACW.PropDef, object> mPropDictionary = new Dictionary<ACW.PropDef, object>();
            ACW.PropDef propDef = new ACW.PropDef();
            ACW.PropDef[] propDefs = conn.WebServiceManager.PropertyService.GetPropertyDefinitionsByEntityClassId("FILE");
            foreach (var item in VaultFileProperties)
            {
                propDef = propDefs.SingleOrDefault(n => n.DispName == item.Key);
                if (propDef != null)
                {
                    //string and datetime types match
                    if (item.Value.GetType().Name == propDef.Typ.ToString())
                    {
                        mPropDictionary.Add(propDef, item.Value);
                    }
                    //number formats are different
                    if (item.Value.GetType().Name == "Decimal" && propDef.Typ.ToString()== "Numeric")
                    {
                        mPropDictionary.Add(propDef, item.Value.ToString());
                    }
                    if (item.Value.GetType().Name == "Boolean" && propDef.Typ.ToString() == "Bool")
                    {
                            mPropDictionary.Add(propDef, item.Value);
                    }
                }
            }

            bool success = mUpdateFileProperties(mFileIt, mPropDictionary);
            if (success)
            {
                if (GetLatestFile == true)
                {
                    string latestFile = GetFileByFullFilePath(VaultFullFileName, false);
                }
                return true;
            }
            return false;

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
                if (IsVaultBasic)
                {
                    settings.OptionsRelationshipGathering.FileRelationshipSettings.VersionGatheringOption = VDF.Vault.Currency.VersionGatheringOption.Actual;
                }
                else
                {
                    settings.OptionsRelationshipGathering.FileRelationshipSettings.ReleaseBiased = true;
                    settings.OptionsRelationshipGathering.FileRelationshipSettings.VersionGatheringOption = VDF.Vault.Currency.VersionGatheringOption.Revision;
                }

                settings.OptionsRelationshipGathering.IncludeLinksSettings.IncludeLinks = false;
                VDF.Vault.Settings.AcquireFilesSettings.AcquireFileResolutionOptions mResOpt = new VDF.Vault.Settings.AcquireFilesSettings.AcquireFileResolutionOptions();
                mResOpt.OverwriteOption = VDF.Vault.Settings.AcquireFilesSettings.AcquireFileResolutionOptions.OverwriteOptions.ForceOverwriteAll;
                mResOpt.SyncWithRemoteSiteSetting = VDF.Vault.Settings.AcquireFilesSettings.SyncWithRemoteSite.Always;
            }

            return settings;
        }

        private VDF.Vault.Currency.Entities.FileIteration CreateFileCopy(ACW.File mSourceFile, string mNewFileName)
        {
            string mExt = String.Format("{0}{1}", ".", (mSourceFile.Name.Split('.')).Last());

            List<long> mIds = new List<long>();
            mIds.Add(mSourceFile.Id);

            ACW.ByteArray mTicket = conn.WebServiceManager.DocumentService.GetDownloadTicketsByFileIds(mIds.ToArray()).First();
            long mTargetFldId = mSourceFile.FolderId;

            ACW.PropWriteResults mResults = new ACW.PropWriteResults();
            byte[] mUploadTicket = conn.WebServiceManager.FilestoreService.CopyFile(mTicket.Bytes, mExt, true, null, out mResults);
            ACW.ByteArray mByteArray = new ACW.ByteArray();
            mByteArray.Bytes = mUploadTicket;

            ACW.File mNewFile = conn.WebServiceManager.DocumentService.AddUploadedFile(mTargetFldId, mNewFileName, "iLogic File Copy from " + mSourceFile.Name, mSourceFile.ModDate, null, null, mSourceFile.FileClass, false, mByteArray);
            VDF.Vault.Currency.Entities.FileIteration mFileIt = new VDF.Vault.Currency.Entities.FileIteration(conn, mNewFile);

            return mFileIt;
        }

        /// <summary>
        /// Update the property dictionary of the given file.
        /// </summary>
        /// <param name="File"></param>
        /// <param name="PropDictionary"></param>
        /// <returns>Returns true on success; returns false in case of failure, e.g., if the file is not available for check out</returns>
        private bool mUpdateFileProperties(ACW.File File, Dictionary<ACW.PropDef, object> PropDictionary)
        {
            try
            {
                ACET.IExplorerUtil mExplUtil = ACET.ExplorerLoader.LoadExplorerUtil(
                                             conn.Server, conn.Vault, conn.UserID, conn.Ticket);
                mExplUtil.UpdateFileProperties(File, PropDictionary);
                return true;
            }
            catch (Exception)
            {
                return false;
            }
        }


        private Image GetThumbnailImage(VDF.Vault.Currency.Entities.FileIteration fileIteration, int Width, int Height)
        {
            try
            {
                VDF.Vault.Currency.Properties.PropertyDefinitionDictionary mPropDefs = conn.PropertyManager.GetPropertyDefinitions(
                                        VDF.Vault.Currency.Entities.EntityClassIds.Files, null, PropertyDefinitionFilter.IncludeSystem);
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
        /// Deprecated. Vault Blog Sample function to convert legacy meta file format and image file format (added to Vault 2013 and later)
        /// </summary>
        /// <param name="value"></param>
        /// <param name="width"></param>
        /// <param name="height"></param>
        /// <returns></returns>
        private Image GetImage(VDF.Vault.Currency.Properties.ThumbnailInfo value, int width, int height)
        {
            byte[] thumbnailRaw = (byte[])value.Image;
            System.Drawing.Image retVal = null;
            using (System.IO.MemoryStream ms = new System.IO.MemoryStream(thumbnailRaw))
            {
                // we don't know what format the image is in, so we try a couple of formats 
                try
                {
                    // try the meta file format 
                    ms.Seek(12, System.IO.SeekOrigin.Begin);
                    System.Drawing.Imaging.Metafile metafile =
                        new System.Drawing.Imaging.Metafile(ms);
                    retVal = metafile.GetThumbnailImage(width, height,
                        new System.Drawing.Image.GetThumbnailImageAbort(GetThumbnailImageAbort),
                        IntPtr.Zero);
                }
                catch
                {
                    // I guess it's not a metafile 
                    retVal = null;
                }


                if (retVal == null)
                {
                    try
                    {
                        // try to stream to System.Drawing.Image 
                        ms.Seek(0, System.IO.SeekOrigin.Begin);
                        System.Drawing.Image rawImage = Image.FromStream(ms, true);
                        retVal = rawImage.GetThumbnailImage(width, height,
                            new System.Drawing.Image.GetThumbnailImageAbort(GetThumbnailImageAbort),
                            IntPtr.Zero);
                    }
                    catch
                    {
                        // not compatible Image object 
                        retVal = null;
                    }
                }
            }
            return retVal;
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

        private bool IsCadFile(System.IO.FileInfo FileInfo)
        {
            //don't add Inventor files except single part files
            List<string> mFileExtensions = new List<string> { ".ipt", ".iam", "ipn", ".idw", ".dwg", ".rvt", ".nwd", ".nwc", ".sldprt", ".sldasm", ".slddrw", ".dgn", ".prt", ".asm", ".drw" };
            if (mFileExtensions.Any(n => FileInfo.Extension == n))
            {
                return true;
            }
            return false;
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
        /// Get File properties as DisplayName/Value map
        /// </summary>
        /// <param name="mFileIt">Autodesk Connectivity Webservice File Version (Iteration) Object</param>
        /// <param name="VaultFileProperties">DisplayName/Value dictionary of file iteration's properties</param>
        private void mGetFileProps(VDF.Vault.Currency.Entities.FileIteration mFileIt, ref Dictionary<string, string> VaultFileProperties)
        {
            ACW.PropDef[] mPropDefs = conn.WebServiceManager.PropertyService.GetPropertyDefinitionsByEntityClassId("FILE");
            ACW.PropInst[] mSourcePropInsts = conn.WebServiceManager.PropertyService.GetPropertiesByEntityIds("FILE", new long[] { mFileIt.EntityIterationId });
            string mPropDispName;
            string mPropVal;
            string mThumbnailDispName = mPropDefs.Where(n => n.SysName == "Thumbnail").FirstOrDefault().DispName;
            foreach (ACW.PropInst mFilePropInst in mSourcePropInsts)
            {
                mPropDispName = mPropDefs.Where(n => n.Id == mFilePropInst.PropDefId).FirstOrDefault().DispName;
                //filter thumbnail property, as iLogic RuleArguments will fail reading it.
                if (mPropDispName != mThumbnailDispName)
                {
                    if (mFilePropInst.Val == null)
                    {
                        mPropVal = "";
                    }
                    else
                    {
                        mPropVal = mFilePropInst.Val.ToString();
                    }
                    VaultFileProperties.Add(mPropDispName, mPropVal);
                }
            }
        }


        /// <summary>
        /// Get Item properties as DisplayName/Value map
        /// </summary>
        /// <param name="item">Autodesk Connectivity Webservice Item object</param>
        /// <param name="VaultItemProperties">DisplayName/Value dictionary of Item version's properties</param>
        private void mGetItemProps(ACW.Item item, ref Dictionary<string, string> VaultItemProperties)
        {
            string mPropDispName = null;
            string mPropVal = null;
            string mThumbnailDispName = null;

            ACW.PropDef[] mItemPropDefs = conn.WebServiceManager.PropertyService.GetPropertyDefinitionsByEntityClassId("ITEM");
            ACW.PropInst[] mItemPropInsts = conn.WebServiceManager.PropertyService.GetPropertiesByEntityIds("ITEM", new long[] { item.Id });
            mThumbnailDispName = mItemPropDefs.Where(n => n.SysName == "Thumbnail").FirstOrDefault().DispName;
            foreach (ACW.PropInst mItemPropInst in mItemPropInsts)
            {
                mPropDispName = mItemPropDefs.Where(n => n.Id == mItemPropInst.PropDefId).FirstOrDefault().DispName;
                if (mPropDispName != mThumbnailDispName)
                {
                    if (mItemPropInst.Val == null)
                    {
                        mPropVal = "";
                    }
                    else
                    {
                        mPropVal = mItemPropInst.Val.ToString();
                    }
                    VaultItemProperties.Add(mPropDispName, mPropVal);
                }
            }
        }

    }

}

