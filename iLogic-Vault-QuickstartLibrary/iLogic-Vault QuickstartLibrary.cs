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
    public class QuickstartiLogicLib :IDisposable
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
        /// Download Vault file using full file path, e.g. "$/Designs/Base.ipt". 
        /// Preset Options: Download Children (recursively) = Enabled, Enforce Overwrite = True
        /// </summary>
        /// <param name="conn">Current Vault Connection</param>
        /// <param name="VaultFullFileName">FullFilePath</param>
        /// <returns>Local path/filename</returns>
        public string mGetFileByFullFileName(VDF.Vault.Currency.Connections.Connection conn, string VaultFullFileName)
        {
            List<string> mFiles = new List<string>();
            mFiles.Add(VaultFullFileName);
            AWS.File[] wsFiles = conn.WebServiceManager.DocumentService.FindLatestFilesByPaths(mFiles.ToArray());
            VDF.Vault.Currency.Entities.FileIteration mFileIt = new VDF.Vault.Currency.Entities.FileIteration(conn,(wsFiles[0]));

            VDF.Vault.Settings.AcquireFilesSettings settings = new VDF.Vault.Settings.AcquireFilesSettings(conn);
            settings.DefaultAcquisitionOption = VCF.Vault.Settings.AcquireFilesSettings.AcquisitionOption.Download;
            settings.OptionsRelationshipGathering.FileRelationshipSettings.IncludeChildren = true;
            settings.OptionsRelationshipGathering.FileRelationshipSettings.RecurseChildren = true;
            settings.OptionsRelationshipGathering.FileRelationshipSettings.VersionGatheringOption = VDF.Vault.Currency.VersionGatheringOption.Latest;
            settings.OptionsRelationshipGathering.IncludeLinksSettings.IncludeLinks = false;
            VCF.Vault.Settings.AcquireFilesSettings.AcquireFileResolutionOptions mResOpt = new VCF.Vault.Settings.AcquireFilesSettings.AcquireFileResolutionOptions();
            mResOpt.OverwriteOption = VCF.Vault.Settings.AcquireFilesSettings.AcquireFileResolutionOptions.OverwriteOptions.ForceOverwriteAll;
            mResOpt.SyncWithRemoteSiteSetting = VCF.Vault.Settings.AcquireFilesSettings.SyncWithRemoteSite.Always;
            settings.AddFileToAcquire(mFileIt, VDF.Vault.Settings.AcquireFilesSettings.AcquisitionOption.Download);
            VDF.Vault.Results.AcquireFilesResults results = conn.FileManager.AcquireFiles(settings);
            if (results != null)
            {
                VDF.Vault.Results.FileAcquisitionResult mFilesDownloaded = results.FileResults.Last();
                return mFilesDownloaded.LocalPath.FullPath.ToString();
            }
            return "FileNotFound";
        }

        /// <summary>
        /// Find file by 1 to many search criteria as property/value pairs. 
        /// Downloads the first file matching all search criterias; include many as required to get the unique file.
        /// Preset Search Operator Options: [Property] is (exactly) [Value]; multiple conditions link up using AND condition.
        /// Preset Download Options: Download Children (recursively) = Enabled, Enforce Overwrite = True
        /// </summary>
        /// <param name="conn">Current Vault Connection</param>
        /// <param name="SearchCriteria">Dictionary of property/value pairs</param>
        /// <returns>Local path/filename</returns>
        public string mGetFilebySearchCriteria(VDF.Vault.Currency.Connections.Connection conn, Dictionary<string,string> SearchCriteria)
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
                    if(i == 0) mSearchCond.SrchRule = AWS.SearchRuleType.May;
                    else mSearchCond.SrchRule = AWS.SearchRuleType.Must;
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
            if (totalResults.Count == 1)
            {
                AWS.File wsFile = totalResults.First<AWS.File>();
                VDF.Vault.Currency.Entities.FileIteration mFileIt = new VDF.Vault.Currency.Entities.FileIteration(conn, (wsFile));

                VDF.Vault.Settings.AcquireFilesSettings settings = new VDF.Vault.Settings.AcquireFilesSettings(conn);
                settings.DefaultAcquisitionOption = VCF.Vault.Settings.AcquireFilesSettings.AcquisitionOption.Download;
                settings.OptionsRelationshipGathering.FileRelationshipSettings.IncludeChildren = true;
                settings.OptionsRelationshipGathering.FileRelationshipSettings.RecurseChildren = true;
                settings.OptionsRelationshipGathering.FileRelationshipSettings.VersionGatheringOption = VDF.Vault.Currency.VersionGatheringOption.Latest;
                settings.OptionsRelationshipGathering.IncludeLinksSettings.IncludeLinks = false;
                VCF.Vault.Settings.AcquireFilesSettings.AcquireFileResolutionOptions mResOpt = new VCF.Vault.Settings.AcquireFilesSettings.AcquireFileResolutionOptions();
                mResOpt.OverwriteOption = VCF.Vault.Settings.AcquireFilesSettings.AcquireFileResolutionOptions.OverwriteOptions.ForceOverwriteAll;
                mResOpt.SyncWithRemoteSiteSetting = VCF.Vault.Settings.AcquireFilesSettings.SyncWithRemoteSite.Always;
                settings.AddFileToAcquire(mFileIt, VDF.Vault.Settings.AcquireFilesSettings.AcquisitionOption.Download);
                VDF.Vault.Results.AcquireFilesResults results = conn.FileManager.AcquireFiles(settings);
                if (results != null)
                {
                    VDF.Vault.Results.FileAcquisitionResult mFilesDownloaded = results.FileResults.Last();
                    return mFilesDownloaded.LocalPath.FullPath.ToString();
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
        /// Downloads all files matching all or any search criterias. Maximum number of files downloaded = Paging size/number.
        /// 
        /// Preset Search Operator Options: [Property] is (exactly) [Value]; multiple conditions link up using AND condition.
        /// Preset Download Options: Download Children (recursively) = Enabled, Enforce Overwrite = True
        /// </summary>
        /// <param name="conn">Current Vault Connection</param>
        /// <param name="SearchCriteria">Dictionary of property/value pairs</param>
        /// <param name="MatchAllCriteria">Switches AND/OR conditions using multiple criterias. Default is false</param>
        /// <returns>Local path/filename</returns>
        public string mGetFilebySearchCriteria(VDF.Vault.Currency.Connections.Connection conn, Dictionary<string, string> SearchCriteria, bool MatchAllCriteria = false)
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
                settings.DefaultAcquisitionOption = VCF.Vault.Settings.AcquireFilesSettings.AcquisitionOption.Download;
                settings.OptionsRelationshipGathering.FileRelationshipSettings.IncludeChildren = true;
                settings.OptionsRelationshipGathering.FileRelationshipSettings.RecurseChildren = true;
                settings.OptionsRelationshipGathering.FileRelationshipSettings.VersionGatheringOption = VDF.Vault.Currency.VersionGatheringOption.Latest;
                settings.OptionsRelationshipGathering.IncludeLinksSettings.IncludeLinks = false;
                VCF.Vault.Settings.AcquireFilesSettings.AcquireFileResolutionOptions mResOpt = new VCF.Vault.Settings.AcquireFilesSettings.AcquireFileResolutionOptions();
                mResOpt.OverwriteOption = VCF.Vault.Settings.AcquireFilesSettings.AcquireFileResolutionOptions.OverwriteOptions.ForceOverwriteAll;
                mResOpt.SyncWithRemoteSiteSetting = VCF.Vault.Settings.AcquireFilesSettings.SyncWithRemoteSite.Always;
                settings.AddFileToAcquire(mFileIt, VDF.Vault.Settings.AcquireFilesSettings.AcquisitionOption.Download);
                VDF.Vault.Results.AcquireFilesResults results = conn.FileManager.AcquireFiles(settings);
                if (results != null)
                {
                    VDF.Vault.Results.FileAcquisitionResult mFilesDownloaded = results.FileResults.Last();
                    return mFilesDownloaded.LocalPath.FullPath.ToString();
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
        /// Returns number of files found, matching the criteria.
        /// 
        /// Preset Search Operator Options: [Property] is (exactly) [Value]; multiple conditions link up using AND/OR condition, depending MatchAllCriteria = True/False
        /// </summary>
        /// <param name="conn">Current Vault Connection</param>
        /// <param name="SearchCriteria">Dictionary of property/value pairs</param>
        /// <param name="MatchAllCriteria">Switches AND/OR conditions using multiple criterias. Default is false</param>
        /// <returns>Number of files found</returns>
        public string[] mCheckFilesExistBySearchCriteria(VDF.Vault.Currency.Connections.Connection conn, Dictionary<string, string> SearchCriteria, bool MatchAllCriteria)
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
    }
}
