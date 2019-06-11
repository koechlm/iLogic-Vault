using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;

using ACET = Autodesk.Connectivity.Explorer.ExtensibilityTools;
using AWS = Autodesk.Connectivity.WebServices;
using VDF = Autodesk.DataManagement.Client.Framework;
using VltBase = Connectivity.Application.VaultBase;

namespace QuickstartiLogicLibrary
{
    /// <summary>
    /// Collection of functions querying and downloading Vault files for iLogic.
    /// </summary>
    public class QuickstartiLogicLib : IDisposable
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
        /// To avoid Vault API specific references, check connection state using the loggedIn property.
        /// </summary>
        private VDF.Vault.Currency.Connections.Connection conn = VltBase.ConnectionManager.Instance.Connection;

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
        /// Deprecated. Returns current Vault connection. Leverage loggedIn property whenever possible. 
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
            mFiles.Add(VaultFullFileName);
            AWS.File[] wsFiles = conn.WebServiceManager.DocumentService.FindLatestFilesByPaths(mFiles.ToArray());
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

            //refine output
            if (results != null)
            {
                try
                {
                    VDF.Vault.Results.FileAcquisitionResult mFilesDownloaded = results.FileResults.Last();
                    return mFilesDownloaded.LocalPath.FullPath.ToString();
                }
                catch (Exception)
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
        /// </summary>
        /// <param name="VaultFullFileName">Vault FullFilePath of source file</param>
        /// <param name="NumberingScheme">Individual scheme name or 'Default'</param>
        /// <param name="InputParams">Optional according scheme definition. User input values in order of scheme configuration</param>
        /// <param name="UpdatePartNumber">Optional. Update Part Number property to match new file name</param>
        /// <param name="CheckOut">Optional. File copy will check-out as default.</param>
        /// <returns>Local path/filename</returns>
        public string GetFileCopyBySourceFileNameAndAutoNumber(string VaultFullFileName, string NumberingScheme, string[] InputParams = null, bool CheckOut = true, bool UpdatePartNumber = true)
        {
            //Get Vault File object
            List<string> mFiles = new List<string>();
            mFiles.Add(VaultFullFileName);
            AWS.File mSourceFile = conn.WebServiceManager.DocumentService.FindLatestFilesByPaths(mFiles.ToArray()).First();

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
                Dictionary<AWS.PropDef, object> mPropDictonary = new Dictionary<AWS.PropDef, object>();

                AWS.PropDef[] propDefs = conn.WebServiceManager.PropertyService.GetPropertyDefinitionsByEntityClassId("FILE");
                AWS.PropDef propDef = propDefs.SingleOrDefault(n => n.SysName == "PartNumber");
                mPropDictonary.Add(propDef, mNewNumber);

                UpdateFileProperties((AWS.File)mFileIt, mPropDictonary);
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

            //refine output
            if (results != null)
                try
                {
                    VDF.Vault.Results.FileAcquisitionResult mFilesDownloaded = results.FileResults.Last();
                    return mFilesDownloaded.LocalPath.FullPath.ToString();
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
        /// </summary>
        /// <param name="VaultFullFileName">Vault FullFilePath of source file</param>
        /// <param name="NewFileNameNoExt">New name without extension</param>
        /// <param name="CheckOut">Optional. File copy will check-out as default.</param>
        /// <param name="UpdatePartNumber">Optional. Update Part Number property to match new file name</param>
        /// <returns>Local path/filename</returns>
        public string GetFileCopyBySourceFileNameAndNewName(string VaultFullFileName, string NewFileNameNoExt, bool CheckOut = true, bool UpdatePartNumber = true )
        {
            //get Vault File object
            List<string> mFiles = new List<string>();
            mFiles.Add(VaultFullFileName);
            AWS.File mSourceFile = conn.WebServiceManager.DocumentService.FindLatestFilesByPaths(mFiles.ToArray()).First();

            if (mSourceFile == null) return null;

            string mNewFileName = String.Format("{0}{1}{2}", NewFileNameNoExt, ".", (mSourceFile.Name).Split('.').Last());

            //create file iteration as copy from source
            VDF.Vault.Currency.Entities.FileIteration mFileIt = CreateFileCopy(mSourceFile, mNewFileName);

            //Optionally update Partnumber property
            if (UpdatePartNumber)
            {
                Dictionary<AWS.PropDef, object> mPropDictonary = new Dictionary<AWS.PropDef, object>();

                AWS.PropDef[] propDefs = conn.WebServiceManager.PropertyService.GetPropertyDefinitionsByEntityClassId("FILE");
                AWS.PropDef propDef = propDefs.SingleOrDefault(n => n.SysName == "PartNumber");
                mPropDictonary.Add(propDef, mNewFileName);

                UpdateFileProperties((AWS.File)mFileIt, mPropDictonary);
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

            //refine output
            if (results != null)
            {
                VDF.Vault.Results.FileAcquisitionResult mFilesDownloaded = results.FileResults.Last();
                return mFilesDownloaded.LocalPath.FullPath.ToString();
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
            AWS.Folder[] mFldr;
            List<long> mFolders = new List<long>();
            if (FoldersSearched != null)
            {
                mFldr = conn.WebServiceManager.DocumentService.FindFoldersByPaths(FoldersSearched);
                foreach (AWS.Folder folder in mFldr)
                {
                    if (folder.Id != -1) mFolders.Add(folder.Id);
                }
            }

            List<String> mFilesFound = new List<string>();
            //combine all search criteria
            List<AWS.SrchCond> mSrchConds = CreateSrchConds(SearchCriteria, MatchAllCriteria);
            List<AWS.File> totalResults = new List<AWS.File>();
            string bookmark = string.Empty;
            AWS.SrchStatus status = null;

            while (status == null || totalResults.Count < status.TotalHits)
            {
                AWS.File[] mSrchResults = conn.WebServiceManager.DocumentService.FindFilesBySearchConditions(
                    mSrchConds.ToArray(), null, mFolders.ToArray(), true, true, ref bookmark, out status);
                if (mSrchResults != null) totalResults.AddRange(mSrchResults);
                else break;
            }
            //if results not empty
            if (totalResults.Count >= 1)
            {
                AWS.File wsFile = totalResults.First<AWS.File>();
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

                //refine output
                if (results != null)
                {
                    try
                    {
                        VDF.Vault.Results.FileAcquisitionResult mFilesDownloaded = results.FileResults.Where(n => n.File.EntityName == totalResults.FirstOrDefault().Name).FirstOrDefault();
                        //return conn.WorkingFoldersManager.GetPathOfFileInWorkingFolder(mFileIt).FullPath.ToString();
                        return mFilesDownloaded.LocalPath.FullPath.ToString();
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
            AWS.Folder[] mFldr;
            List<long> mFolders = new List<long>();
            if (FoldersSearched != null)
            {
                mFldr = conn.WebServiceManager.DocumentService.FindFoldersByPaths(FoldersSearched);
                foreach (AWS.Folder folder in mFldr)
                {
                    if (folder.Id != -1) mFolders.Add(folder.Id);
                }
            }

            List<VDF.Vault.Currency.Entities.FileIteration> mFilesFound = new List<VDF.Vault.Currency.Entities.FileIteration>();
            List<String> mFilesDownloaded = new List<string>();
            //combine all search criteria
            List<AWS.SrchCond> mSrchConds = CreateSrchConds(SearchCriteria, MatchAllCriteria);
            List<AWS.File> totalResults = new List<AWS.File>();
            string bookmark = string.Empty;
            AWS.SrchStatus status = null;

            while (status == null || totalResults.Count < status.TotalHits)
            {
                AWS.File[] mSrchResults = conn.WebServiceManager.DocumentService.FindFilesBySearchConditions(
                    mSrchConds.ToArray(), null, mFolders.ToArray(), true, true, ref bookmark, out status);
                if (mSrchResults != null) totalResults.AddRange(mSrchResults);
                else break;
            }
            //if results not empty
            if (totalResults.Count >= 1)
            {
                //build download options including DefaultAcquisitionOptions
                VDF.Vault.Settings.AcquireFilesSettings settings = CreateAcquireSettings(false);
                foreach (AWS.File wsFile in totalResults)
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
                    foreach (AWS.File wsFile in totalResults)
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

                    //foreach (var item in results.FileResults)
                    //{
                    //    if (totalResults.Any(n=>n.Name == item.File.EntityName))
                    //    {
                    //        mFilesFound.Add(item.LocalPath.FullPath.ToString());
                    //    }
                    //}
                    //return mFilesFound;
                }
                return null;
            }
            else
            {
                return null;
            }
        }


        /// <summary>
        /// Find 1 to many file(s) by 1 to many search criteria as property/value pairs. 
        /// Downloads first file matching all or any search criterias.
        /// Preset Search Operator Options: [Property] is (exactly) [Value]; multiple conditions link up using AND condition.
        /// Preset Download Options: Download Children (recursively) = Enabled, Enforce Overwrite = True
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
            //FoldersSearched: Inventor files are expected in IPJ registered path's only. In case of null use these:
            AWS.Folder[] mFldr;
            List<long> mFolders = new List<long>();
            if (FoldersSearched != null)
            {
                mFldr = conn.WebServiceManager.DocumentService.FindFoldersByPaths(FoldersSearched);
                foreach (AWS.Folder folder in mFldr)
                {
                    if (folder.Id != -1) mFolders.Add(folder.Id);
                }
            }

            List<String> mFilesFound = new List<string>();
            //combine all search criteria
            List<AWS.SrchCond> mSrchConds = CreateSrchConds(SearchCriteria, MatchAllCriteria);
            List<AWS.File> totalResults = new List<AWS.File>();
            string bookmark = string.Empty;
            AWS.SrchStatus status = null;

            while (status == null || totalResults.Count < status.TotalHits)
            {
                AWS.File[] mSrchResults = conn.WebServiceManager.DocumentService.FindFilesBySearchConditions(
                    mSrchConds.ToArray(), null, mFolders.ToArray(), true, true, ref bookmark, out status);
                if (mSrchResults != null) totalResults.AddRange(mSrchResults);
                else break;
            }
            //if results not empty
            if (totalResults.Count >= 1)
            {
                AWS.File mSourceFile = totalResults.First<AWS.File>();
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
                    Dictionary<AWS.PropDef, object> mPropDictonary = new Dictionary<AWS.PropDef, object>();

                    AWS.PropDef[] propDefs = conn.WebServiceManager.PropertyService.GetPropertyDefinitionsByEntityClassId("FILE");
                    AWS.PropDef propDef = propDefs.SingleOrDefault(n => n.SysName == "PartNumber");
                    mPropDictonary.Add(propDef, mNewNumber);

                    UpdateFileProperties((AWS.File)mFileIt, mPropDictonary);
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

                //refine output
                if (results != null)
                {
                    try
                    {
                        VDF.Vault.Results.FileAcquisitionResult mFilesDownloaded = results.FileResults.Last();
                        return mFilesDownloaded.LocalPath.FullPath.ToString();
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
        /// Downloads first file matching all or any search criterias.
        /// Preset Search Operator Options: [Property] is (exactly) [Value]; multiple conditions link up using AND condition.
        /// Preset Download Options: Download Children (recursively) = Enabled, Enforce Overwrite = True
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
            //FoldersSearched: Inventor files are expected in IPJ registered path's only. In case of null use these:
            AWS.Folder[] mFldr;
            List<long> mFolders = new List<long>();
            if (FoldersSearched != null)
            {
                mFldr = conn.WebServiceManager.DocumentService.FindFoldersByPaths(FoldersSearched);
                foreach (AWS.Folder folder in mFldr)
                {
                    if (folder.Id != -1) mFolders.Add(folder.Id);
                }
            }

            List<String> mFilesFound = new List<string>();
            //combine all search criteria
            List<AWS.SrchCond> mSrchConds = CreateSrchConds(SearchCriteria, MatchAllCriteria);
            List<AWS.File> totalResults = new List<AWS.File>();
            string bookmark = string.Empty;
            AWS.SrchStatus status = null;

            while (status == null || totalResults.Count < status.TotalHits)
            {
                AWS.File[] mSrchResults = conn.WebServiceManager.DocumentService.FindFilesBySearchConditions(
                    mSrchConds.ToArray(), null, mFolders.ToArray(), true, true, ref bookmark, out status);
                if (mSrchResults != null) totalResults.AddRange(mSrchResults);
                else break;
            }
            //if results not empty
            if (totalResults.Count >= 1)
            {
                AWS.File mSourceFile = totalResults.First<AWS.File>();
                if (mSourceFile == null) return null;

                string mNewFileName = String.Format("{0}{1}{2}", NewFileNameNoExt, ".", (mSourceFile.Name).Split('.').Last());
                VDF.Vault.Currency.Entities.FileIteration mFileIt = CreateFileCopy(mSourceFile, mNewFileName);

                //Optionally pdate Partnumber property
                if (UpdatePartNumber)
                {
                    Dictionary<AWS.PropDef, object> mPropDictonary = new Dictionary<AWS.PropDef, object>();

                    AWS.PropDef[] propDefs = conn.WebServiceManager.PropertyService.GetPropertyDefinitionsByEntityClassId("FILE");
                    AWS.PropDef propDef = propDefs.SingleOrDefault(n => n.SysName == "PartNumber");
                    mPropDictonary.Add(propDef, mNewFileName);
                    
                    UpdateFileProperties((AWS.File)mFileIt, mPropDictonary);
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

                //refine output
                if (results != null)
                {
                    try
                    {
                        VDF.Vault.Results.FileAcquisitionResult mFilesDownloaded = results.FileResults.Last();
                        return mFilesDownloaded.LocalPath.FullPath.ToString();
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
        /// Returns array of file names found, matching the criteria.
        /// 
        /// Preset Search Operator Options: [Property] is (exactly) [Value]; multiple conditions link up using AND/OR condition, depending MatchAllCriteria = True/False
        /// </summary>
        /// <param name="SearchCriteria">Dictionary of property/value pairs</param>
        /// <param name="MatchAllCriteria">Optional. Switches AND/OR conditions using multiple criterias. Default is true</param>
        /// <param name="FoldersSearched">Optional. Limit search scope to given folder path(s).</param>
        /// <returns>Array of file names found</returns>
        public IList<string> CheckFilesExistBySearchCriteria(Dictionary<string, string> SearchCriteria, bool MatchAllCriteria = true, string[] FoldersSearched = null)
        {
            //FoldersSearched: Inventor files are expected in IPJ registered path's only. In case of null use these:
            AWS.Folder[] mFldr;
            List<long> mFolders = new List<long>();
            if (FoldersSearched != null)
            {
                mFldr = conn.WebServiceManager.DocumentService.FindFoldersByPaths(FoldersSearched);
                foreach (AWS.Folder folder in mFldr)
                {
                    if (folder.Id != -1) mFolders.Add(folder.Id);
                }
            }

            List<String> mFilesFound = new List<string>();
            //combine all search criteria
            List<AWS.SrchCond> mSrchConds = CreateSrchConds(SearchCriteria, MatchAllCriteria);
            List<AWS.File> totalResults = new List<AWS.File>();
            string bookmark = string.Empty;
            AWS.SrchStatus status = null;

            while (status == null || totalResults.Count < status.TotalHits)
            {
                AWS.File[] mSrchResults = conn.WebServiceManager.DocumentService.FindFilesBySearchConditions(
                     mSrchConds.ToArray(), null, mFolders.ToArray(), true, true, ref bookmark, out status);
                if (mSrchResults != null) totalResults.AddRange(mSrchResults);
                else break;
            }
            //if results not empty
            if (totalResults.Count >= 1)
            {
                foreach (AWS.File wsFile in totalResults)
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
        /// Create single file number by scheme name and optional input parameters
        /// </summary>
        /// <param name="mSchmName">Name of individual Numbering Scheme or "Default" for pre-set scheme</param>
        /// <param name="mSchmPrms">User input parameter in order of scheme configuration</param>
        /// <returns>new number</returns>
        public string GetNewNumber(string mSchmName, string[] mSchmPrms = null)
        {
            AWS.NumSchm NmngSchm = null;
            if (mSchmPrms == null)
            {
                mSchmPrms = Array.Empty<string>();
            }

            try
            {
                if (mSchmName == "Default")
                {
                    NmngSchm = conn.WebServiceManager.NumberingService.GetNumberingSchemes("FILE", AWS.NumSchmType.Activated).First(n => n.IsDflt == true);
                }
                else
                {
                    NmngSchm = conn.WebServiceManager.NumberingService.GetNumberingSchemes("FILE", AWS.NumSchmType.Activated).First(n => n.Name == mSchmName);
                }
                return conn.WebServiceManager.DocumentService.GenerateFileNumber(NmngSchm.SchmID, null);
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


        private List<AWS.SrchCond> CreateSrchConds(Dictionary<string, string> SearchCriteria, bool MatchAllCriteria)
        {
            AWS.PropDef[] mFilePropDefs = conn.WebServiceManager.PropertyService.GetPropertyDefinitionsByEntityClassId("FILE");
            //iterate mSearchcriteria to get property definitions and build AWS search criteria
            List<AWS.SrchCond> mSrchConds = new List<AWS.SrchCond>();
            int i = 0;
            foreach (var item in SearchCriteria)
            {
                AWS.PropDef mFilePropDef = mFilePropDefs.Single(n => n.DispName == item.Key);
                AWS.SrchCond mSearchCond = new AWS.SrchCond();
                {
                    mSearchCond.PropDefId = mFilePropDef.Id;
                    mSearchCond.PropTyp = AWS.PropertySearchType.SingleProperty;
                    mSearchCond.SrchOper = 1; //equals
                    if (i == 0) mSearchCond.SrchRule = AWS.SearchRuleType.Must;
                    else
                    {
                        if (MatchAllCriteria) mSearchCond.SrchRule = AWS.SearchRuleType.Must;
                        else mSearchCond.SrchRule = AWS.SearchRuleType.May;
                    }
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
                settings.OptionsRelationshipGathering.FileRelationshipSettings.ReleaseBiased = true;
                settings.OptionsRelationshipGathering.FileRelationshipSettings.VersionGatheringOption = VDF.Vault.Currency.VersionGatheringOption.Revision;
                settings.OptionsRelationshipGathering.IncludeLinksSettings.IncludeLinks = false;
                VDF.Vault.Settings.AcquireFilesSettings.AcquireFileResolutionOptions mResOpt = new VDF.Vault.Settings.AcquireFilesSettings.AcquireFileResolutionOptions();
                mResOpt.OverwriteOption = VDF.Vault.Settings.AcquireFilesSettings.AcquireFileResolutionOptions.OverwriteOptions.ForceOverwriteAll;
                mResOpt.SyncWithRemoteSiteSetting = VDF.Vault.Settings.AcquireFilesSettings.SyncWithRemoteSite.Always;
            }

            return settings;
        }


        VDF.Vault.Currency.Entities.FileIteration CreateFileCopy(AWS.File mSourceFile, string mNewFileName)
        {
            string mExt = String.Format("{0}{1}", ".", (mSourceFile.Name.Split('.')).Last());

            List<long> mIds = new List<long>();
            mIds.Add(mSourceFile.Id);

            AWS.ByteArray mTicket = conn.WebServiceManager.DocumentService.GetDownloadTicketsByFileIds(mIds.ToArray()).First();
            long mTargetFldId = mSourceFile.FolderId;

            AWS.PropWriteResults mResults = new AWS.PropWriteResults();
            byte[] mUploadTicket = conn.WebServiceManager.FilestoreService.CopyFile(mTicket.Bytes, mExt, true, null, out mResults);
            AWS.ByteArray mByteArray = new AWS.ByteArray();
            mByteArray.Bytes = mUploadTicket;

            AWS.File mNewFile = conn.WebServiceManager.DocumentService.AddUploadedFile(mTargetFldId, mNewFileName, "iLogic File Copy from " + mSourceFile.Name, mSourceFile.ModDate, null, null, mSourceFile.FileClass, false, mByteArray);
            VDF.Vault.Currency.Entities.FileIteration mFileIt = new VDF.Vault.Currency.Entities.FileIteration(conn, mNewFile);

            return mFileIt;
        }


        /// <summary>
        /// Prepared for future use.
        /// </summary>
        /// <param name="mFile"></param>
        /// <param name="mPropDictonary"></param>
        /// <returns></returns>
        private bool UpdateFileProperties(AWS.File mFile, Dictionary<AWS.PropDef, object> mPropDictonary)
        {
            try
            {
                ACET.IExplorerUtil mExplUtil = Autodesk.Connectivity.Explorer.ExtensibilityTools.ExplorerLoader.LoadExplorerUtil(
                                            conn.Server, conn.Vault, conn.UserID, conn.Ticket);
                mExplUtil.UpdateFileProperties(mFile, mPropDictonary);
                return true;
            }
            catch
            {
                return false;
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
                AWS.File[] wsFiles = conn.WebServiceManager.DocumentService.FindLatestFilesByPaths(mFiles.ToArray());
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
            AWS.Folder[] mFldr;
            List<long> mFolders = new List<long>();
            if (FoldersSearched != null)
            {
                mFldr = conn.WebServiceManager.DocumentService.FindFoldersByPaths(FoldersSearched);
                foreach (AWS.Folder folder in mFldr)
                {
                    if (folder.Id != -1) mFolders.Add(folder.Id);
                }
            }

            List<String> mFilesFound = new List<string>();
            //combine all search criteria
            List<AWS.SrchCond> mSrchConds = CreateSrchConds(SearchCriteria, MatchAllCriteria);
            List<AWS.File> totalResults = new List<AWS.File>();
            string bookmark = string.Empty;
            AWS.SrchStatus status = null;

            while (status == null || totalResults.Count < status.TotalHits)
            {
                AWS.File[] mSrchResults = conn.WebServiceManager.DocumentService.FindFilesBySearchConditions(
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
                    AWS.File wsFile = totalResults.First<AWS.File>();
                    VDF.Vault.Currency.Entities.FileIteration mFileIt = new VDF.Vault.Currency.Entities.FileIteration(conn, (wsFile));

                    Image image = GetThumbnailImage(mFileIt, Height, Width);
                    if (image != null)
                    {
                        AWS.Folder mParentFldr = conn.WebServiceManager.DocumentService.GetFolderById(wsFile.FolderId);
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
        /// <param name="VaultFullFileName">Full Vault source file path of format '$\...\*.*'</param>
        /// <param name="Width">Optional. Image pixel size</param>
        /// <param name="Height">Optional. Image pixel size.</param>
        /// <returns>System.Drawing.Image object</returns>
        public Image GetThumbnailImageByFullSourceFilePath(string VaultFullFileName, int Width = 300, int Height = 300)
        {
            List<string> mFiles = new List<string>();
            mFiles.Add(VaultFullFileName);
            try
            {
                AWS.File[] wsFiles = conn.WebServiceManager.DocumentService.FindLatestFilesByPaths(mFiles.ToArray());
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
            AWS.Folder[] mFldr;
            List<long> mFolders = new List<long>();
            if (FoldersSearched != null)
            {
                mFldr = conn.WebServiceManager.DocumentService.FindFoldersByPaths(FoldersSearched);
                foreach (AWS.Folder folder in mFldr)
                {
                    if (folder.Id != -1) mFolders.Add(folder.Id);
                }
            }

            List<String> mFilesFound = new List<string>();
            //combine all search criteria
            List<AWS.SrchCond> mSrchConds = CreateSrchConds(SearchCriteria, MatchAllCriteria);
            List<AWS.File> totalResults = new List<AWS.File>();
            string bookmark = string.Empty;
            AWS.SrchStatus status = null;

            while (status == null || totalResults.Count < status.TotalHits)
            {
                AWS.File[] mSrchResults = conn.WebServiceManager.DocumentService.FindFilesBySearchConditions(
                    mSrchConds.ToArray(), null, mFolders.ToArray(), true, true, ref bookmark, out status);
                if (mSrchResults != null) totalResults.AddRange(mSrchResults);
                else break;
            }
            //if results not empty
            if (totalResults.Count >= 1)
            {
                try
                {
                    AWS.File wsFile = totalResults.First<AWS.File>();
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
                        System.IntPtr.Zero);
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
                        System.Drawing.Image rawImage = System.Drawing.Image.FromStream(ms, true);
                        retVal = rawImage.GetThumbnailImage(width, height,
                            new System.Drawing.Image.GetThumbnailImageAbort(GetThumbnailImageAbort),
                            System.IntPtr.Zero);
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
    }
}
