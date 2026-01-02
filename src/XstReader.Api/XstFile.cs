// Project site: https://github.com/iluvadev/XstReader
//
// Based on the great work of Dijji. 
// Original project: https://github.com/dijji/XstReader
//
// Issues: https://github.com/iluvadev/XstReader/issues
// License (Ms-PL): https://github.com/iluvadev/XstReader/blob/master/license.md
//
// Copyright (c) 2021, iluvadev, and released under Ms-PL License.
// Copyright (c) 2016, Dijji, and released under Ms-PL.  This can be found in the root of this distribution. 

using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Text;
using XstReader.Common;
using XstReader.ElementProperties;

namespace XstReader
{
    // See: https://docs.microsoft.com/en-us/openspecs/exchange_server_protocols/ms-oxcmsg/7fd7ec40-deec-4c06-9493-1bc06b349682


    // The code here implements the messaging layer, which depends on and invokes the NDP and LTP layers

    /// <summary>
    /// Main handling for xst (.ost and .pst) files 
    /// </summary>
    public class XstFile : XstElement, IDisposable
    {
        private NDB _Ndb;
        internal new NDB Ndb => _Ndb ?? (_Ndb = new NDB(this));

        private LTP _Ltp;
        internal new LTP Ltp => _Ltp ?? (_Ltp = new LTP(Ndb));

        private static readonly Encoding[] PasswordEncodings = new[] { Encoding.Unicode, Encoding.UTF8, Encoding.ASCII };
        private const UInt16 PstPasswordPropertyTag = 0x67ff;
        private string _password;
        private bool _passwordValidated;
        private XstPropertySet _MessageStoreProperties;

        private string _FileName = null;
        /// <summary>
        /// FileName of the .pst or .ost file to read
        /// </summary>
        public string FileName { get => _FileName; set => SetFileName(value); }

        /// <summary>
        /// Optional password used to unlock message store access when the PST enforces it.
        /// </summary>
        public string Password
        {
            get => _password;
            set
            {
                if (_password == value)
                    return;
                _password = value;
                _passwordValidated = false;
            }
        }
        private void SetFileName(string fileName)
        {
            _FileName = fileName;
            ClearContents();
        }

        private FileStream _ReadStream = null;
        internal FileStream ReadStream
        {
            get => _ReadStream ?? (_ReadStream = new FileStream(FileName, FileMode.Open, FileAccess.Read));
        }
        internal object StreamLock { get; } = new object();

        private XstPropertySet MessageStoreProperties
        {
            get
            {
                if (_MessageStoreProperties == null)
                {
                    var messageStoreNid = new NID((uint)EnidSpecial.NID_MESSAGE_STORE);
                    _MessageStoreProperties = new XstPropertySet(
                        () => Ltp.ReadAllProperties(messageStoreNid),
                        tag => Ltp.ReadProperty(messageStoreNid, tag),
                        tag => Ltp.ContainsProperty(messageStoreNid, tag));
                }

                return _MessageStoreProperties;
            }
        }

        private XstFolder _RootFolder = null;
        /// <summary>
        /// The Root Folder of the XstFile. (Loaded when needed)
        /// </summary>
        public XstFolder RootFolder
        {
            get
            {
                EnsurePasswordUnlocked();
                return _RootFolder ?? (_RootFolder = new XstFolder(this, new NID(EnidSpecial.NID_ROOT_FOLDER)));
            }
        }

        /// <summary>
        /// The Path of this Element
        /// </summary>
        [DisplayName("Path")]
        [Category("General")]
        [Description(@"The Path of this Element")]
        public override string Path => System.IO.Path.GetFileName(this.FileName);

        /// <summary>
        /// The Parents of this Element
        /// </summary>
        [Browsable(false)]
        public override XstElement Parent => null;


        #region Ctor
        /// <summary>
        /// Ctor
        /// </summary>
        /// <param name="fileName">The .pst or .ost file to open</param>
        /// <param name="password">Optional password used to validate access when the PST message store requires it.</param>
        public XstFile(string fileName, string password = null) : base(XstElementType.File)
        {
            Password = password;
            FileName = fileName;
        }
        #endregion Ctor

        private void EnsurePasswordUnlocked()
        {
            if (_passwordValidated)
                return;

            var passwordProperty = MessageStoreProperties?.Get(PstPasswordPropertyTag);
            if (!(passwordProperty?.Value is int storedCrcValue))
            {
                _passwordValidated = true;
                return;
            }

            var storedCrc = unchecked((uint)storedCrcValue);
            if (storedCrc == 0)
            {
                _passwordValidated = true;
                return;
            }

            if (string.IsNullOrEmpty(Password))
                throw new XstException("The PST file is password protected. Set XstFile.Password before accessing its contents.");

            if (!PasswordMatches(storedCrc, Password))
                throw new XstException($"The supplied PST password is incorrect. (Expected CRC 0x{storedCrc:X8})");

            _passwordValidated = true;
        }

        private static bool PasswordMatches(uint storedCrc, string password)
        {
            if (string.IsNullOrEmpty(password))
                return false;

            var upper = password.ToUpperInvariant();
            bool differsByCase = !string.Equals(password, upper, StringComparison.Ordinal);

            foreach (var encoding in PasswordEncodings)
            {
                if (TryMatch(storedCrc, password, encoding))
                    return true;

                if (differsByCase && TryMatch(storedCrc, upper, encoding))
                    return true;
            }

            return false;
        }

        private static bool TryMatch(uint storedCrc, string value, Encoding encoding)
        {
            var bytes = encoding.GetBytes(value);
            if (Crc32.Compute(bytes, 0, bytes.Length) == storedCrc)
                return true;

            var terminator = encoding.GetBytes("\0");
            if (terminator.Length > 0)
            {
                var withNull = new byte[bytes.Length + terminator.Length];
                Buffer.BlockCopy(bytes, 0, withNull, 0, bytes.Length);
                Buffer.BlockCopy(terminator, 0, withNull, bytes.Length, terminator.Length);
                if (Crc32.Compute(withNull, 0, withNull.Length) == storedCrc)
                    return true;
            }

            return false;
        }

        private void ClearStream()
        {
            if (_ReadStream != null)
            {
                _ReadStream.Close();
                _ReadStream.Dispose();
                _ReadStream = null;
            }
        }

        /// <summary>
        /// Clears information and memory used in RootFolder
        /// </summary>
        private void ClearRootFolder()
        {
            if (_RootFolder != null)
            {
                _RootFolder.ClearContents();
                _RootFolder = null;
            }
        }

        /// <summary>
        /// Clears all information and memory used by the object
        /// </summary>
        public override void ClearContents()
        {
            ClearStream();
            ClearRootFolder();

            _Ndb = null;
            _Ltp = null;
            _MessageStoreProperties?.ClearContents();
            _MessageStoreProperties = null;
            _passwordValidated = false;
        }

        /// <summary>
        /// Disposes memory used by the object
        /// </summary>
        public void Dispose()
        {
            ClearContents();
        }

        /// <summary>
        /// Gets the String representation of the object
        /// </summary>
        /// <returns></returns>
        public override string ToString()
        {
            return System.IO.Path.GetFileName(FileName ?? "");
        }

        private protected override IEnumerable<XstProperty> LoadProperties()
        {
            return new XstProperty[0];
        }

        private protected override XstProperty LoadProperty(PropertyCanonicalName tag)
        {
            return null;
        }

        private protected override bool CheckProperty(PropertyCanonicalName tag)
        {
            return false;
        }
    }
}
