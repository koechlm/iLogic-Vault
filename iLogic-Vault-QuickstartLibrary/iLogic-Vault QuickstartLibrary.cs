using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;

using VCF = Autodesk.DataManagement.Client.Framework;

using AWS = Autodesk.Connectivity.WebServices;
using VDF = Autodesk.DataManagement.Client.Framework;
using Autodesk.DataManagement.Client.Framework.Currency;
using VB = Connectivity.Application.VaultBase;

namespace QuickstartiLogicLibrary
{
    public class QuickstartiLogicLib : IDisposable
    {
        public void Dispose()
        {
            //do clean up here if required
        }

        /// <summary>
        /// Returns current Vault connection; required for any iLogic-Vault communication. 
        /// Null value returned if user is not logged in.
        /// </summary>
        /// <returns>Vault Connection</returns>
        public VDF.Vault.Currency.Connections.Connection mGetVaultConn()
        {
            VDF.Vault.Currency.Connections.Connection mConn = VB.ConnectionManager.Instance.Connection;
            if (mConn != null)
            {
                return mConn;
            }
            return null;
        }

        /// <summary>
        /// Downloads Vault file using full file path, e.g. "$/Designs/Base.ipt". Returns full file name in local working folder (download enforces override, if local file exists),
        /// returns "FileNotFound if file does not exist at indicated location.
        /// Preset Options: Download Children (recursively) = Enabled, Enforce Overwrite = True
        /// </summary>
        /// <param name="conn">Current Vault Connection</param>
        /// <param name="VaultFullFileName">FullFilePath</param>
        /// <param name="CheckOut">Optional. File downloaded does NOT check-out as default.</param>
        /// <returns>Local path/filename or error statement "FileNotFound"</returns>
        public string mGetFileByFullFileName(VDF.Vault.Currency.Connections.Connection conn, string VaultFullFileName, bool CheckOut = false)
        {
            List<string> mFiles = new List<string>();
            mFiles.Add(VaultFullFileName);
            AWS.File[] wsFiles = conn.WebServiceManager.DocumentService.FindLatestFilesByPaths(mFiles.ToArray());
            VDF.Vault.Currency.Entities.FileIteration mFileIt = new VDF.Vault.Currency.Entities.FileIteration(conn, (wsFiles[0]));

            VDF.Vault.Settings.AcquireFilesSettings settings = new VDF.Vault.Settings.AcquireFilesSettings(conn);
            if (CheckOut)
            {
                settings.DefaultAcquisitionOption = VCF.Vault.Settings.AcquireFilesSettings.AcquisitionOption.Checkout;
            }
            else
            {
                settings.DefaultAcquisitionOption = VCF.Vault.Settings.AcquireFilesSettings.AcquisitionOption.Download;
            }
            settings.OptionsRelationshipGathering.FileRelationshipSettings.IncludeChildren = true;
            settings.OptionsRelationshipGathering.FileRelationshipSettings.RecurseChildren = true;
            settings.OptionsRelationshipGathering.FileRelationshipSettings.VersionGatheringOption = VDF.Vault.Currency.VersionGatheringOption.Latest;
            settings.OptionsRelationshipGathering.IncludeLinksSettings.IncludeLinks = false;
            VCF.Vault.Settings.AcquireFilesSettings.AcquireFileResolutionOptions mResOpt = new VCF.Vault.Settings.AcquireFilesSettings.AcquireFileResolutionOptions();
            mResOpt.OverwriteOption = VCF.Vault.Settings.AcquireFilesSettings.AcquireFileResolutionOptions.OverwriteOptions.ForceOverwriteAll;
            mResOpt.SyncWithRemoteSiteSetting = VCF.Vault.Settings.AcquireFilesSettings.SyncWithRemoteSite.Always;
            settings.AddFileToAcquire(mFileIt, settings.DefaultAcquisitionOption);
            VDF.Vault.Results.AcquireFilesResults results = conn.FileManager.AcquireFiles(settings);
            if (results != null)
            {
                try
                {
                    VDF.Vault.Results.FileAcquisitionResult mFilesDownloaded = results.FileResults.Last();
                    return mFilesDownloaded.LocalPath.FullPath.ToString();
                }
                catch (Exception)
                {
                    return "FileFoundButDownloadFailed";
                }
            }
            return "FileNotFound";
        }

        /// <summary>
        /// Copy Vault file on file server and download using full file path, e.g. "$/Designs/Base.ipt".
        /// Create new file name using default or named numbering scheme.
        /// Preset Options: Download Children (recursively) = Enabled, Enforce Overwrite = True
        /// </summary>
        /// <param name="conn">Current Vault Connection</param>
        /// <param name="VaultFullFileName">Vault FullFilePath of source file</param>
        /// <param name="NumberingScheme">Individual scheme name or 'Default'</param>
        /// <param name="InputParams">Optional according scheme definition. User input values in order of scheme configuration</param>
        /// <param name="CheckOut">Optional. File copy will check-out as default.</param>
        /// <returns>Local path/filename or error statement of "SourceFileNotFound" or "GetNumberFailed" or "VaultFileCopyFailure"</returns>
        public string mGetFileCopyBySourceFileNameAndAutoNumber(VDF.Vault.Currency.Connections.Connection conn, string VaultFullFileName, string NumberingScheme, string[] InputParams = null, bool CheckOut = true)
        {
            List<string> mFiles = new List<string>();
            mFiles.Add(VaultFullFileName);
            AWS.File mSourceFile = conn.WebServiceManager.DocumentService.FindLatestFilesByPaths(mFiles.ToArray()).First();

            if (mSourceFile == null) return "SourceFileNotFound";

            string mExt = String.Format("{0}{1}", ".", (VaultFullFileName.Split('.')).Last());

            List<long> mIds = new List<long>();
            mIds.Add(mSourceFile.Id);

            AWS.ByteArray mTicket = conn.WebServiceManager.DocumentService.GetDownloadTicketsByFileIds(mIds.ToArray()).First();
            long mTargetFldId = mSourceFile.FolderId;

            AWS.PropWriteResults mResults = new AWS.PropWriteResults();
            byte[] mUploadTicket = conn.WebServiceManager.FilestoreService.CopyFile(mTicket.Bytes, mExt, true, null, out mResults);
            AWS.ByteArray mByteArray = new AWS.ByteArray();
            mByteArray.Bytes = mUploadTicket;

            string mNewNumber = mGetNewNumber(conn, NumberingScheme, InputParams);

            if (mNewNumber == "GetNumberFailed") return "GetNumberFailed";

            string mNewFileName = String.Format("{0}{1}{2}", mNewNumber, ".", (mSourceFile.Name).Split('.').Last());

            AWS.File mNewFile = conn.WebServiceManager.DocumentService.AddUploadedFile(mTargetFldId, mNewFileName, "iLogic File Copy from " + VaultFullFileName, mSourceFile.ModDate, null, null, mSourceFile.FileClass, false, mByteArray);

            VDF.Vault.Currency.Entities.FileIteration mFileIt = new VDF.Vault.Currency.Entities.FileIteration(conn, mNewFile);

            VDF.Vault.Settings.AcquireFilesSettings settings = new VDF.Vault.Settings.AcquireFilesSettings(conn);
            if (CheckOut)
            {
                settings.DefaultAcquisitionOption = VCF.Vault.Settings.AcquireFilesSettings.AcquisitionOption.Checkout;
            }
            else
            {
                settings.DefaultAcquisitionOption = VCF.Vault.Settings.AcquireFilesSettings.AcquisitionOption.Download;
            }
            settings.OptionsRelationshipGathering.FileRelationshipSettings.IncludeChildren = true;
            settings.OptionsRelationshipGathering.FileRelationshipSettings.RecurseChildren = true;
            settings.OptionsRelationshipGathering.FileRelationshipSettings.VersionGatheringOption = VDF.Vault.Currency.VersionGatheringOption.Latest;
            settings.OptionsRelationshipGathering.IncludeLinksSettings.IncludeLinks = false;

            VCF.Vault.Settings.AcquireFilesSettings.AcquireFileResolutionOptions mResOpt = new VCF.Vault.Settings.AcquireFilesSettings.AcquireFileResolutionOptions();
            mResOpt.OverwriteOption = VCF.Vault.Settings.AcquireFilesSettings.AcquireFileResolutionOptions.OverwriteOptions.ForceOverwriteAll;
            mResOpt.SyncWithRemoteSiteSetting = VCF.Vault.Settings.AcquireFilesSettings.SyncWithRemoteSite.Always;

            settings.AddFileToAcquire(mFileIt, settings.DefaultAcquisitionOption);
            VDF.Vault.Results.AcquireFilesResults results = conn.FileManager.AcquireFiles(settings);
            if (results != null)
                try
                {
                    VDF.Vault.Results.FileAcquisitionResult mFilesDownloaded = results.FileResults.Last();
                    return mFilesDownloaded.LocalPath.FullPath.ToString();
                }
                catch (Exception)
                {
                    return "FileCopiedButDownloadFailed";
                }
            return "VaultFileCopyFailure";
        }


        /// <summary>
        /// Copy Vault file on file server and download using full file path, e.g. "$/Designs/Base.ipt".
        /// Create new file name re-using source file's extension and new file name variable.
        /// Preset Options: Download Children (recursively) = Enabled, Enforce Overwrite = True
        /// </summary>
        /// <param name="conn">Current Vault Connection</param>
        /// <param name="VaultFullFileName">Vault FullFilePath of source file</param>
        /// <param name="NewFileNameNoExt">New name without extension</param>
        /// <param name="CheckOut">Optional. File copy will check-out as default.</param>
        /// <returns>Local path/filename or error statement of "SourceFileNotFound" or "VaultFileCopyFailure"</returns>
        public string mGetFileCopyBySourceFileNameAndNewName(VDF.Vault.Currency.Connections.Connection conn, string VaultFullFileName, string NewFileNameNoExt, bool CheckOut = true)
        {
            List<string> mFiles = new List<string>();
            mFiles.Add(VaultFullFileName);
            AWS.File mSourceFile = conn.WebServiceManager.DocumentService.FindLatestFilesByPaths(mFiles.ToArray()).First();

            if (mSourceFile == null) return "SourceFileNotFound";

            string mExt = String.Format("{0}{1}", ".", (VaultFullFileName.Split('.')).Last());

            List<long> mIds = new List<long>();
            mIds.Add(mSourceFile.Id);

            AWS.ByteArray mTicket = conn.WebServiceManager.DocumentService.GetDownloadTicketsByFileIds(mIds.ToArray()).First();
            long mTargetFldId = mSourceFile.FolderId;

            AWS.PropWriteResults mResults = new AWS.PropWriteResults();
            byte[] mUploadTicket = conn.WebServiceManager.FilestoreService.CopyFile(mTicket.Bytes, mExt, true, null, out mResults);
            AWS.ByteArray mByteArray = new AWS.ByteArray();
            mByteArray.Bytes = mUploadTicket;

            string mNewFileName = String.Format("{0}{1}{2}", NewFileNameNoExt, ".", (mSourceFile.Name).Split('.').Last());

            AWS.File mNewFile = conn.WebServiceManager.DocumentService.AddUploadedFile(mTargetFldId, mNewFileName, "iLogic File Copy from " + VaultFullFileName, mSourceFile.ModDate, null, null, mSourceFile.FileClass, false, mByteArray);

            VDF.Vault.Currency.Entities.FileIteration mFileIt = new VDF.Vault.Currency.Entities.FileIteration(conn, mNewFile);

            VDF.Vault.Settings.AcquireFilesSettings settings = new VDF.Vault.Settings.AcquireFilesSettings(conn);
            if (CheckOut)
            {
                settings.DefaultAcquisitionOption = VCF.Vault.Settings.AcquireFilesSettings.AcquisitionOption.Checkout;
            }
            else
            {
                settings.DefaultAcquisitionOption = VCF.Vault.Settings.AcquireFilesSettings.AcquisitionOption.Download;
            }
            settings.OptionsRelationshipGathering.FileRelationshipSettings.IncludeChildren = true;
            settings.OptionsRelationshipGathering.FileRelationshipSettings.RecurseChildren = true;
            settings.OptionsRelationshipGathering.FileRelationshipSettings.VersionGatheringOption = VDF.Vault.Currency.VersionGatheringOption.Latest;
            settings.OptionsRelationshipGathering.IncludeLinksSettings.IncludeLinks = false;

            VCF.Vault.Settings.AcquireFilesSettings.AcquireFileResolutionOptions mResOpt = new VCF.Vault.Settings.AcquireFilesSettings.AcquireFileResolutionOptions();
            mResOpt.OverwriteOption = VCF.Vault.Settings.AcquireFilesSettings.AcquireFileResolutionOptions.OverwriteOptions.ForceOverwriteAll;
            mResOpt.SyncWithRemoteSiteSetting = VCF.Vault.Settings.AcquireFilesSettings.SyncWithRemoteSite.Always;

            settings.AddFileToAcquire(mFileIt, settings.DefaultAcquisitionOption);
            VDF.Vault.Results.AcquireFilesResults results = conn.FileManager.AcquireFiles(settings);
            if (results != null)
                if (results != null)
                {
                    VDF.Vault.Results.FileAcquisitionResult mFilesDownloaded = results.FileResults.Last();
                    return mFilesDownloaded.LocalPath.FullPath.ToString();
                }
            return "VaultFileCopyFailed";
        }


        /// <summary>
        /// Find 1 to many file(s) by 1 to many search criteria as property/value pairs. 
        /// Downloads first file matching all or any search criterias. 
        /// Preset Search Operator Options: [Property] is (exactly) [Value]; multiple conditions link up using AND condition.
        /// Preset Download Options: Download Children (recursively) = Enabled, Enforce Overwrite = True
        /// </summary>
        /// <param name="conn">Current Vault Connection</param>
        /// <param name="SearchCriteria">Dictionary of property/value pairs</param>
        /// <param name="MatchAllCriteria">Optional. Switches AND/OR conditions using multiple criterias. Default is false</param>
        /// <param name="CheckOut">Optional. File downloaded does NOT check-out as default</param>
        /// <returns>Local path/filename or error statement "FileNotFound"</returns>
        public string mGetFilebySearchCriteria(VDF.Vault.Currency.Connections.Connection conn, Dictionary<string, string> SearchCriteria, bool MatchAllCriteria = false, bool CheckOut = false)
        {
            AWS.PropDef[] mFilePropDefs = conn.WebServiceManager.PropertyService.GetPropertyDefinitionsByEntityClassId("FILE");
            //iterate mSearchcriteria to get property definitions and build AWS search criteria
            List<AWS.SrchCond> mSrchConds = new List<AWS.SrchCond>();
            int i = 0;
            List<AWS.File> totalResults = new List<AWS.File>();
            foreach (var item in SearchCriteria)
            {
                AWS.PropDef mFilePropDef = mFilePropDefs.Single(n => n.DispName == item.Key);
                AWS.SrchCond mSearchCond = new AWS.SrchCond();
                {
                    mSearchCond.PropDefId = mFilePropDef.Id;
                    mSearchCond.PropTyp = AWS.PropertySearchType.SingleProperty;
                    mSearchCond.SrchOper = 1; //equals
                    if (i == 0) mSearchCond.SrchRule = AWS.SearchRuleType.May;
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
            string bookmark = string.Empty;
            AWS.SrchStatus status = null;

            while (status == null || totalResults.Count < status.TotalHits)
            {
                AWS.File[] mSrchResults = conn.WebServiceManager.DocumentService.FindFilesBySearchConditions(
                    mSrchConds.ToArray(), null, null, false, true, ref bookmark, out status);
                if (mSrchResults != null) totalResults.AddRange(mSrchResults);
                else break;
            }
            //if results not empty
            if (totalResults.Count >= 1)
            {
                AWS.File wsFile = totalResults.First<AWS.File>();
                VDF.Vault.Currency.Entities.FileIteration mFileIt = new VDF.Vault.Currency.Entities.FileIteration(conn, (wsFile));

                VDF.Vault.Settings.AcquireFilesSettings settings = new VDF.Vault.Settings.AcquireFilesSettings(conn);
                if (CheckOut)
                {
                    settings.DefaultAcquisitionOption = VCF.Vault.Settings.AcquireFilesSettings.AcquisitionOption.Checkout;
                }
                else
                {
                    settings.DefaultAcquisitionOption = VCF.Vault.Settings.AcquireFilesSettings.AcquisitionOption.Download;
                }
                settings.OptionsRelationshipGathering.FileRelationshipSettings.IncludeChildren = true;
                settings.OptionsRelationshipGathering.FileRelationshipSettings.RecurseChildren = true;
                settings.OptionsRelationshipGathering.FileRelationshipSettings.VersionGatheringOption = VDF.Vault.Currency.VersionGatheringOption.Latest;
                settings.OptionsRelationshipGathering.IncludeLinksSettings.IncludeLinks = false;
                VCF.Vault.Settings.AcquireFilesSettings.AcquireFileResolutionOptions mResOpt = new VCF.Vault.Settings.AcquireFilesSettings.AcquireFileResolutionOptions();
                mResOpt.OverwriteOption = VCF.Vault.Settings.AcquireFilesSettings.AcquireFileResolutionOptions.OverwriteOptions.ForceOverwriteAll;
                mResOpt.SyncWithRemoteSiteSetting = VCF.Vault.Settings.AcquireFilesSettings.SyncWithRemoteSite.Always;
                settings.AddFileToAcquire(mFileIt, settings.DefaultAcquisitionOption);
                VDF.Vault.Results.AcquireFilesResults results = conn.FileManager.AcquireFiles(settings);
                if (results != null)
                {
                    try
                    {
                        VDF.Vault.Results.FileAcquisitionResult mFilesDownloaded = results.FileResults.Last();
                        return mFilesDownloaded.LocalPath.FullPath.ToString();
                    }
                    catch (Exception)
                    {
                        return "FileFoundButDownloadFailed";
                    }
                }
                else
                {
                    return "FileNotFound";
                }
            }
            else
            {
                return "FileNotFound";
            }
        }

        /// <summary>
        /// Find 1 to many file(s) by 1 to many search criteria as property/value pairs. 
        /// Downloads first file matching all or any search criterias.
        /// Preset Search Operator Options: [Property] is (exactly) [Value]; multiple conditions link up using AND condition.
        /// Preset Download Options: Download Children (recursively) = Enabled, Enforce Overwrite = True
        /// </summary>
        /// <param name="conn">Current Vault Connection</param>
        /// <param name="SearchCriteria">Dictionary of property/value pairs</param>        
        /// <param name="NumberingScheme">Individual scheme name or 'Default'</param>
        /// <param name="MatchAllCriteria">Optional. Switches AND/OR conditions using multiple criterias. Default is false</param>
        /// <param name="InputParams">Optional according scheme definition. User input values in order of scheme configuration</param>
        /// <param name="CheckOut">Optional. File copy will check-out as default.</param>
        /// <returns>Local path/filename or error statement "SourceFileNotFound" or "GetNumberFailed" or "VaultFileCopyFailure"</returns>
        public string mGetFileCopyBySourceFileSearchAndAutoNumber(VDF.Vault.Currency.Connections.Connection conn, Dictionary<string, string> SearchCriteria, string NumberingScheme, bool MatchAllCriteria = false, string[] InputParams = null, bool CheckOut = true)
        {
            AWS.PropDef[] mFilePropDefs = conn.WebServiceManager.PropertyService.GetPropertyDefinitionsByEntityClassId("FILE");
            //iterate mSearchcriteria to get property definitions and build AWS search criteria
            List<AWS.SrchCond> mSrchConds = new List<AWS.SrchCond>();
            int i = 0;
            List<AWS.File> totalResults = new List<AWS.File>();
            foreach (var item in SearchCriteria)
            {
                AWS.PropDef mFilePropDef = mFilePropDefs.Single(n => n.DispName == item.Key);
                AWS.SrchCond mSearchCond = new AWS.SrchCond();
                {
                    mSearchCond.PropDefId = mFilePropDef.Id;
                    mSearchCond.PropTyp = AWS.PropertySearchType.SingleProperty;
                    mSearchCond.SrchOper = 1; //equals
                    if (i == 0) mSearchCond.SrchRule = AWS.SearchRuleType.May;
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
            string bookmark = string.Empty;
            AWS.SrchStatus status = null;

            while (status == null || totalResults.Count < status.TotalHits)
            {
                AWS.File[] mSrchResults = conn.WebServiceManager.DocumentService.FindFilesBySearchConditions(
                    mSrchConds.ToArray(), null, null, false, true, ref bookmark, out status);
                if (mSrchResults != null) totalResults.AddRange(mSrchResults);
                else break;
            }
            //if results not empty
            if (totalResults.Count >= 1)
            {
                AWS.File mSourceFile = totalResults.First<AWS.File>();
                if (mSourceFile == null) return "SourceFileNotFound";

                string mExt = String.Format("{0}{1}", ".", (mSourceFile.Name.Split('.')).Last());

                List<long> mIds = new List<long>();
                mIds.Add(mSourceFile.Id);

                AWS.ByteArray mTicket = conn.WebServiceManager.DocumentService.GetDownloadTicketsByFileIds(mIds.ToArray()).First();
                long mTargetFldId = mSourceFile.FolderId;

                AWS.PropWriteResults mResults = new AWS.PropWriteResults();
                byte[] mUploadTicket = conn.WebServiceManager.FilestoreService.CopyFile(mTicket.Bytes, mExt, true, null, out mResults);
                AWS.ByteArray mByteArray = new AWS.ByteArray();
                mByteArray.Bytes = mUploadTicket;

                string mNewNumber = mGetNewNumber(conn, NumberingScheme, InputParams);

                if (mNewNumber == "GetNumberFailed") return "GetNumberFailed";

                string mNewFileName = String.Format("{0}{1}{2}", mNewNumber, ".", (mSourceFile.Name).Split('.').Last());

                AWS.File mNewFile = conn.WebServiceManager.DocumentService.AddUploadedFile(mTargetFldId, mNewFileName, "iLogic File Copy from " + mSourceFile.Name, mSourceFile.ModDate, null, null, mSourceFile.FileClass, false, mByteArray);

                VDF.Vault.Currency.Entities.FileIteration mFileIt = new VDF.Vault.Currency.Entities.FileIteration(conn, mNewFile);

                VDF.Vault.Settings.AcquireFilesSettings settings = new VDF.Vault.Settings.AcquireFilesSettings(conn);
                if (CheckOut)
                {
                    settings.DefaultAcquisitionOption = VCF.Vault.Settings.AcquireFilesSettings.AcquisitionOption.Checkout;
                }
                else
                {
                    settings.DefaultAcquisitionOption = VCF.Vault.Settings.AcquireFilesSettings.AcquisitionOption.Download;
                }
                settings.OptionsRelationshipGathering.FileRelationshipSettings.IncludeChildren = true;
                settings.OptionsRelationshipGathering.FileRelationshipSettings.RecurseChildren = true;
                settings.OptionsRelationshipGathering.FileRelationshipSettings.VersionGatheringOption = VDF.Vault.Currency.VersionGatheringOption.Latest;
                settings.OptionsRelationshipGathering.IncludeLinksSettings.IncludeLinks = false;
                VCF.Vault.Settings.AcquireFilesSettings.AcquireFileResolutionOptions mResOpt = new VCF.Vault.Settings.AcquireFilesSettings.AcquireFileResolutionOptions();
                mResOpt.OverwriteOption = VCF.Vault.Settings.AcquireFilesSettings.AcquireFileResolutionOptions.OverwriteOptions.ForceOverwriteAll;
                mResOpt.SyncWithRemoteSiteSetting = VCF.Vault.Settings.AcquireFilesSettings.SyncWithRemoteSite.Always;
                settings.AddFileToAcquire(mFileIt, settings.DefaultAcquisitionOption);
                VDF.Vault.Results.AcquireFilesResults results = conn.FileManager.AcquireFiles(settings);
                if (results != null)
                {
                    try
                    {
                        VDF.Vault.Results.FileAcquisitionResult mFilesDownloaded = results.FileResults.Last();
                        return mFilesDownloaded.LocalPath.FullPath.ToString();
                    }
                    catch (Exception)
                    {
                        return "FileCopiedButDownloadFailed";
                    }
                }
                else
                {
                    return "FileNotFound";
                }
            }
            else
            {
                return "FileNotFound";
            }
        }

        /// <summary>
        /// Find 1 to many file(s) by 1 to many search criteria as property/value pairs. 
        /// Downloads first file matching all or any search criterias.
        /// Preset Search Operator Options: [Property] is (exactly) [Value]; multiple conditions link up using AND condition.
        /// Preset Download Options: Download Children (recursively) = Enabled, Enforce Overwrite = True
        /// </summary>
        /// <param name="conn">Current Vault Connection</param>
        /// <param name="SearchCriteria">Dictionary of property/value pairs</param>        
        /// <param name="NewFileNameNoExt">New name without extension</param>
        /// <param name="MatchAllCriteria">Optional. Switches AND/OR conditions using multiple criterias. Default is false</param>
        /// <param name="CheckOut">Optional. File copy will check-out as default.</param>
        /// <returns>Local path/filename or error statement "SourceFileNotFound" or "GetNumberFailed" or "VaultFileCopyFailure"</returns>
        public string mGetFileCopyBySourceFileSearchAndNewName(VDF.Vault.Currency.Connections.Connection conn, Dictionary<string, string> SearchCriteria, string NewFileNameNoExt, bool MatchAllCriteria = false, bool CheckOut = true)
        {
            AWS.PropDef[] mFilePropDefs = conn.WebServiceManager.PropertyService.GetPropertyDefinitionsByEntityClassId("FILE");
            //iterate mSearchcriteria to get property definitions and build AWS search criteria
            List<AWS.SrchCond> mSrchConds = new List<AWS.SrchCond>();
            int i = 0;
            List<AWS.File> totalResults = new List<AWS.File>();
            foreach (var item in SearchCriteria)
            {
                AWS.PropDef mFilePropDef = mFilePropDefs.Single(n => n.DispName == item.Key);
                AWS.SrchCond mSearchCond = new AWS.SrchCond();
                {
                    mSearchCond.PropDefId = mFilePropDef.Id;
                    mSearchCond.PropTyp = AWS.PropertySearchType.SingleProperty;
                    mSearchCond.SrchOper = 1; //equals
                    if (i == 0) mSearchCond.SrchRule = AWS.SearchRuleType.May;
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
            string bookmark = string.Empty;
            AWS.SrchStatus status = null;

            while (status == null || totalResults.Count < status.TotalHits)
            {
                AWS.File[] mSrchResults = conn.WebServiceManager.DocumentService.FindFilesBySearchConditions(
                    mSrchConds.ToArray(), null, null, false, true, ref bookmark, out status);
                if (mSrchResults != null) totalResults.AddRange(mSrchResults);
                else break;
            }
            //if results not empty
            if (totalResults.Count >= 1)
            {
                AWS.File mSourceFile = totalResults.First<AWS.File>();
                if (mSourceFile == null) return "SourceFileNotFound";

                string mExt = String.Format("{0}{1}", ".", (mSourceFile.Name.Split('.')).Last());

                List<long> mIds = new List<long>();
                mIds.Add(mSourceFile.Id);

                AWS.ByteArray mTicket = conn.WebServiceManager.DocumentService.GetDownloadTicketsByFileIds(mIds.ToArray()).First();
                long mTargetFldId = mSourceFile.FolderId;

                AWS.PropWriteResults mResults = new AWS.PropWriteResults();
                byte[] mUploadTicket = conn.WebServiceManager.FilestoreService.CopyFile(mTicket.Bytes, mExt, true, null, out mResults);
                AWS.ByteArray mByteArray = new AWS.ByteArray();
                mByteArray.Bytes = mUploadTicket;

                string mNewFileName = String.Format("{0}{1}{2}", NewFileNameNoExt, ".", (mSourceFile.Name).Split('.').Last());

                AWS.File mNewFile = conn.WebServiceManager.DocumentService.AddUploadedFile(mTargetFldId, mNewFileName, "iLogic File Copy from " + mSourceFile.Name, mSourceFile.ModDate, null, null, mSourceFile.FileClass, false, mByteArray);

                VDF.Vault.Currency.Entities.FileIteration mFileIt = new VDF.Vault.Currency.Entities.FileIteration(conn, mNewFile);

                VDF.Vault.Settings.AcquireFilesSettings settings = new VDF.Vault.Settings.AcquireFilesSettings(conn);
                if (CheckOut)
                {
                    settings.DefaultAcquisitionOption = VCF.Vault.Settings.AcquireFilesSettings.AcquisitionOption.Checkout;
                }
                else
                {
                    settings.DefaultAcquisitionOption = VCF.Vault.Settings.AcquireFilesSettings.AcquisitionOption.Download;
                }
                settings.OptionsRelationshipGathering.FileRelationshipSettings.IncludeChildren = true;
                settings.OptionsRelationshipGathering.FileRelationshipSettings.RecurseChildren = true;
                settings.OptionsRelationshipGathering.FileRelationshipSettings.VersionGatheringOption = VDF.Vault.Currency.VersionGatheringOption.Latest;
                settings.OptionsRelationshipGathering.IncludeLinksSettings.IncludeLinks = false;
                VCF.Vault.Settings.AcquireFilesSettings.AcquireFileResolutionOptions mResOpt = new VCF.Vault.Settings.AcquireFilesSettings.AcquireFileResolutionOptions();
                mResOpt.OverwriteOption = VCF.Vault.Settings.AcquireFilesSettings.AcquireFileResolutionOptions.OverwriteOptions.ForceOverwriteAll;
                mResOpt.SyncWithRemoteSiteSetting = VCF.Vault.Settings.AcquireFilesSettings.SyncWithRemoteSite.Always;
                settings.AddFileToAcquire(mFileIt, settings.DefaultAcquisitionOption);
                VDF.Vault.Results.AcquireFilesResults results = conn.FileManager.AcquireFiles(settings);
                if (results != null)
                {
                    try
                    {
                        VDF.Vault.Results.FileAcquisitionResult mFilesDownloaded = results.FileResults.Last();
                        return mFilesDownloaded.LocalPath.FullPath.ToString();
                    }
                    catch (Exception)
                    {
                        return "FileCopiedButDownloadFailed";
                    }
                }
                else
                {
                    return "FileNotFound";
                }
            }
            else
            {
                return "FileNotFound";
            }
        }


        /// <summary>
        /// Find 1 to many file(s) by 1 to many search criteria as property/value pairs. 
        /// Returns array of file names found, matching the criteria.
        /// 
        /// Preset Search Operator Options: [Property] is (exactly) [Value]; multiple conditions link up using AND/OR condition, depending MatchAllCriteria = True/False
        /// </summary>
        /// <param name="conn">Current Vault Connection</param>
        /// <param name="SearchCriteria">Dictionary of property/value pairs</param>
        /// <param name="MatchAllCriteria">Optional. Switches AND/OR conditions using multiple criterias. Default is false</param>
        /// <returns>Array of file names found</returns>
        public string[] mCheckFilesExistBySearchCriteria(VDF.Vault.Currency.Connections.Connection conn, Dictionary<string, string> SearchCriteria, bool MatchAllCriteria = false)
        {
            List<String> mFilesFound = new List<string>();
            AWS.PropDef[] mFilePropDefs = conn.WebServiceManager.PropertyService.GetPropertyDefinitionsByEntityClassId("FILE");
            //iterate mSearchcriteria to get property definitions and build AWS search criteria
            List<AWS.SrchCond> mSrchConds = new List<AWS.SrchCond>();
            int i = 0;
            List<AWS.File> totalResults = new List<AWS.File>();
            foreach (var item in SearchCriteria)
            {
                AWS.PropDef mFilePropDef = mFilePropDefs.Single(n => n.DispName == item.Key);
                AWS.SrchCond mSearchCond = new AWS.SrchCond();
                {
                    mSearchCond.PropDefId = mFilePropDef.Id;
                    mSearchCond.PropTyp = AWS.PropertySearchType.SingleProperty;
                    mSearchCond.SrchOper = 1; //equals
                    if (i == 0) mSearchCond.SrchRule = AWS.SearchRuleType.May;
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
            string bookmark = string.Empty;
            AWS.SrchStatus status = null;

            while (status == null || totalResults.Count < status.TotalHits)
            {
                AWS.File[] mSrchResults = conn.WebServiceManager.DocumentService.FindFilesBySearchConditions(
                    mSrchConds.ToArray(), null, null, false, true, ref bookmark, out status);
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
                return mFilesFound.ToArray();
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
        /// <param name="conn">Current Vault Connection</param>
        /// <param name="SearchCriteria">Dictionary of property/value pairs</param>
        /// <param name="MatchAllCriteria">Optional. Switches AND/OR conditions using multiple criterias. Default is false</param>
        /// <param name="CheckOut">Optional. Downloaded files will NOT check-out as default.</param>
        /// <returns>Array of file names found</returns>
        public string[] mGetMultipleFilesBySearch(VDF.Vault.Currency.Connections.Connection conn, Dictionary<string, string> SearchCriteria, bool MatchAllCriteria = false, bool CheckOut = false)
        {
            List<String> mFilesFound = new List<string>();
            AWS.PropDef[] mFilePropDefs = conn.WebServiceManager.PropertyService.GetPropertyDefinitionsByEntityClassId("FILE");
            //iterate mSearchcriteria to get property definitions and build AWS search criteria
            List<AWS.SrchCond> mSrchConds = new List<AWS.SrchCond>();
            int i = 0;
            List<AWS.File> totalResults = new List<AWS.File>();
            foreach (var item in SearchCriteria)
            {
                AWS.PropDef mFilePropDef = mFilePropDefs.Single(n => n.DispName == item.Key);
                AWS.SrchCond mSearchCond = new AWS.SrchCond();
                {
                    mSearchCond.PropDefId = mFilePropDef.Id;
                    mSearchCond.PropTyp = AWS.PropertySearchType.SingleProperty;
                    mSearchCond.SrchOper = 1; //equals
                    if (i == 0) mSearchCond.SrchRule = AWS.SearchRuleType.May;
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
            string bookmark = string.Empty;
            AWS.SrchStatus status = null;

            while (status == null || totalResults.Count < status.TotalHits)
            {
                AWS.File[] mSrchResults = conn.WebServiceManager.DocumentService.FindFilesBySearchConditions(
                    mSrchConds.ToArray(), null, null, false, true, ref bookmark, out status);
                if (mSrchResults != null) totalResults.AddRange(mSrchResults);
                else break;
            }
            //if results not empty
            if (totalResults.Count >= 1)
            {
                VDF.Vault.Settings.AcquireFilesSettings settings = new VDF.Vault.Settings.AcquireFilesSettings(conn);
                if (CheckOut)
                {
                    settings.DefaultAcquisitionOption = VCF.Vault.Settings.AcquireFilesSettings.AcquisitionOption.Checkout;
                }
                else
                {
                    settings.DefaultAcquisitionOption = VCF.Vault.Settings.AcquireFilesSettings.AcquisitionOption.Download;
                }
                settings.OptionsRelationshipGathering.FileRelationshipSettings.IncludeChildren = true;
                settings.OptionsRelationshipGathering.FileRelationshipSettings.RecurseChildren = true;
                settings.OptionsRelationshipGathering.FileRelationshipSettings.VersionGatheringOption = VDF.Vault.Currency.VersionGatheringOption.Latest;
                settings.OptionsRelationshipGathering.IncludeLinksSettings.IncludeLinks = false;
                VCF.Vault.Settings.AcquireFilesSettings.AcquireFileResolutionOptions mResOpt = new VCF.Vault.Settings.AcquireFilesSettings.AcquireFileResolutionOptions();
                mResOpt.OverwriteOption = VCF.Vault.Settings.AcquireFilesSettings.AcquireFileResolutionOptions.OverwriteOptions.ForceOverwriteAll;
                mResOpt.SyncWithRemoteSiteSetting = VCF.Vault.Settings.AcquireFilesSettings.SyncWithRemoteSite.Always;

                foreach (AWS.File wsFile in totalResults)
                {
                    VDF.Vault.Currency.Entities.FileIteration mFileIt = new VDF.Vault.Currency.Entities.FileIteration(conn, wsFile);

                    settings.AddFileToAcquire(mFileIt, settings.DefaultAcquisitionOption);
                }

                VDF.Vault.Results.AcquireFilesResults results = conn.FileManager.AcquireFiles(settings);
                if (results.FileResults != null)
                {
                    foreach (var item in results.FileResults)
                    {
                        mFilesFound.Add(item.LocalPath.FullPath.ToString());
                    }
                    return mFilesFound.ToArray();
                }
                return null;
            }
            else
            {
                return null;
            }
        }


        /// <summary>
        /// Create single file number by scheme name and optional input parameters
        /// </summary>
        /// <param name="conn">Current Vault connection</param>
        /// <param name="mSchmName">Name of individual Numbering Scheme or "Default" for pre-set scheme</param>
        /// <param name="mSchmPrms">User input parameter in order of scheme configuration</param>
        /// <returns>new number or error message "GetNumberFailed"</returns>
        public string mGetNewNumber(VDF.Vault.Currency.Connections.Connection conn, string mSchmName, string[] mSchmPrms = null)
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
                return "GetNumberFailed";
            }
        }

        /// <summary>
        /// Copies a local file to a new name. 
        /// The source file's location and extension are captured and apply to the copy.
        /// Use Check-In (iLogic) command for adding the new file to Vault.
        /// </summary>
        /// <param name="mFullFileName">File name including full path</param>
        /// <param name="mNewNameNoExtension">The new target name of the copied file. Path and extension will transfer from the source file.</param>
        /// <returns>Local path/filename or error statement "LocalFileCopyFailed"</returns>
        public string mCopyLocalFile(string mFullFileName, string mNewNameNoExtension)
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
                return "LocalFileCopyFailed";
            }
        }
    }
}
