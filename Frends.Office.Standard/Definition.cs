#pragma warning disable 1591

using System.ComponentModel;
using System.ComponentModel.DataAnnotations;

namespace Frends.Office.Standard
{
    /// <summary>
    /// Input for excel file writers.
    /// </summary>
    public class WriteExcelFileInput
    {
        /// <summary>
        /// Input csv string.
        /// </summary>
        [DefaultValue("\"1;2;3\\r\\na;b;c\"")]
        [DisplayFormat(DataFormatString = "Expression")]
        public string StringInput { get; set; }

        /// <summary>
        /// Determines what character will be used for splitting based on cell in csv. Deafult is ';'.
        /// </summary>
        [DefaultValue("';'")]
        [DisplayFormat(DataFormatString = "Expression")]
        public char CellDelimiter { get; set; }

        /// <summary>
        /// Determines what string will be used for splitting lines. Default is "\r\n".
        /// </summary>
        [DefaultValue("\"\\r\\n\"")]
        [DisplayFormat(DataFormatString = "Expression")]
        public string LineDelimiter { get; set; }

        /// <summary>
        /// Full path of the target file to be written. File format should be .xlsx, e.g. FileName.xlsx
        /// </summary>
        [DefaultValue(@"c:\temp\file.xlsx")]
        [DisplayFormat(DataFormatString = "Text")]
        public string TargetPath { get; set; }
    }

    /// <summary>
    /// Input for file writers.
    /// </summary>
    public class WriteWordFileInput
    {
        /// <summary>
        /// Input string data.
        /// </summary>
        [DefaultValue("\"Test input\\r\\nNew line\"")]
        [DisplayFormat(DataFormatString = "Expression")]
        public string StringInput { get; set; }

        /// <summary>
        /// Determines what string will be used for splitting lines. Default is "\r\n".
        /// </summary>
        [DefaultValue("\"\\r\\n\"")]
        [DisplayFormat(DataFormatString = "Expression")]
        public string LineDelimiter { get; set; }

        /// <summary>
        /// Full path of the target file to be written. File format should be .docx, e.g. FileName.docx
        /// </summary>
        [DefaultValue(@"c:\file.docx")]
        [DisplayFormat(DataFormatString = "Text")]
        public string TargetPath { get; set; }
    }
    /// <summary>
    /// Input data class for exporting files to Sharepoint.
    /// </summary>
    public class ExportFileInput
    {
        /// <summary>
        /// Full path of the target file to be written, e.g. c:\FileName.xlsx
        /// </summary>
        [DefaultValue("@\"c:\\temp\\file.xlsx\"")]
        [DisplayFormat(DataFormatString = "Expression")]
        public string SourceFilePath { get; set; }

        /// <summary>
        /// Target folder path on Sharepoint.
        /// </summary>
        [DisplayFormat(DataFormatString = "Expression")]
        public string TargetFolderPath { get; set; }
    }
    /// <summary>
    /// Input data class for authenticating to Sharepoint.
    /// </summary>
    public class SharepointAuthentication
    {
        /// <summary>
        /// Azure Active Directory Registered APP Client ID.
        /// </summary>
        [DisplayFormat(DataFormatString = "Expression")]
        public string ClientID { get; set; }

        /// <summary>
        /// Azure Active Directory Registered APP Client Secret.
        /// </summary>
        [DisplayFormat(DataFormatString = "Expression")]
        public string ClientSecret { get; set; }

        /// <summary>
        /// Azure Active Directory tenant ID.
        /// </summary>
        [DisplayFormat(DataFormatString = "Expression")]
        public string TenantID { get; set; }

        /// <summary>
        /// Sharepoint Site ID - retrievable from Microsoft API once the site is created.
        /// </summary>
        [DisplayFormat(DataFormatString = "Expression")]
        public string SiteID { get; set; }
    }
}
