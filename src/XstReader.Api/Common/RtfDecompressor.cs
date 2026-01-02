// Project site: https://github.com/iluvadev/XstReader
//
// Based on the great work of Dijji. 
// Original project: https://github.com/dijji/XstReader
//
// Issues: https://github.com/iluvadev/XstReader/issues
// License (Ms-PL): https://github.com/iluvadev/XstReader/blob/master/license.md
//
// Copyright (c) 2016, Dijji, and released under Ms-PL.  This can be found in the root of this distribution. 

using System;
using System.Collections;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;

namespace XstReader.Common
{
    // Implementation of the RTF decompression algorithm specified in [MS-OXRTFCP]
    // This is a port of the VB code at http://www.vbforums.com/showthread.php?669883-NET-3-5-RtfDecompressor-Decompress-RTF-From-Outlook-And-Exchange-Server
    internal class RtfDecompressor
    {
        //fields
        private byte[] InitialDictionary;
        //constants
        private const int HeaderLength = 0x10;
        private const int CircularDictionaryMaxLength = 0x1000;

        //constructors
        public RtfDecompressor()
        {
            //initialize dictionary, must be this exact string
            var builder = new StringBuilder();

            builder.Append(@"{\rtf1\ansi\mac\deff0\deftab720{\fonttbl;}");
            builder.Append(@"{\f0\fnil \froman \fswiss \fmodern \fscript ");
            builder.Append(@"\fdecor MS Sans SerifSymbolArialTimes New RomanCourier{\colortbl\red0\green0\blue0");
            builder.Append("\r\n");
            builder.Append(@"\par \pard\plain\f0\fs20\b\i\u\tab\tx");

            InitialDictionary = Encoding.ASCII.GetBytes(builder.ToString()); //2.1.2.1
        }

        //methods

        /// <summary>
        /// Decompresses an RTF <see cref="Stream">Stream</see> and returns the decompressed stream as an array of bytes.
        /// </summary>
        /// <param name="stream">The <see cref="Stream">Stream</see> to decompress.</param>
        /// <param name="enforceCrc">True to enforce a CRC check; otherwise, false to ignore CRC checking.</param>
        /// <exception cref="System.IndexOutOfRangeException">Thrown when the stream reaches a corrupt or unpredicted state.</exception>
        /// <returns>The decompressed byte stream.</returns>
        public MemoryStream Decompress(Stream stream, bool enforceCrc = false)
        {
            if (stream.CanRead)
            {
                var buffer = new byte[stream.Length - 1];
                stream.Read(buffer, 0, (int)stream.Length);
                return Decompress(buffer, enforceCrc);
            }
            return null;
        }

        /// <summary>
        /// Decompresses an RTF byte stream and returns the decompressed stream as an array of bytes.
        /// </summary>
        /// <param name="data">The compressed stream to decompress.</param>
        /// <param name="enforceCrc">True to enforce a CRC check; otherwise, false to ignore CRC checking.</param>
        /// <exception cref="System.IndexOutOfRangeException">Thrown when the stream reaches a corrupt or unpredicted state.</exception>
        /// <returns>The decompressed byte stream.</returns>
        public MemoryStream Decompress(byte[] data, bool enforceCrc = false)
        {
            //2.2.3.1.2
            var header = Map.MapType<RtfHeader>(data);
            var initialLength = InitialDictionary.Length;

            switch (header.compType)
            {

                case (UInt32)CompressionTypes.UnCompressed:

                    //data is uncompressed, this is very rare
                    //Should the header be excluded from what we return?
                    return new MemoryStream(data, HeaderLength, data.Length - HeaderLength);

                case (UInt32)CompressionTypes.Compressed:

                    //2.2.3
                    if (enforceCrc)
                    {
                        var crc = CalculateCrc(data, HeaderLength);
                        if (crc != header.crc)
                            throw new XstException("Input stream is corrupt: CRC did not match");
                    }

                    byte[] dictionary = new byte[CircularDictionaryMaxLength];
                    var destination = new MemoryStream((int)header.rawSize);
                    // Initialise the dictionary
                    Array.Copy(InitialDictionary, 0, dictionary, 0, initialLength);
                    int dictionaryWrite = initialLength;
                    int dictionaryEnd = initialLength;

                    try
                    {
                        for (int i = HeaderLength; i < data.Length;)
                        {
                            var control = new BitArray(new byte[] { data[i] });
                            int offset = 1;

                            for (int j = 0; j < control.Length; j++)
                            {
                                if (!control[j])
                                {
                                    //literal bit
                                    destination.WriteByte(data[i + offset]);
                                    dictionary[dictionaryWrite++] = (data[i + offset]);
                                    if (dictionaryWrite > dictionaryEnd)
                                        dictionaryEnd = Math.Min(dictionaryWrite, CircularDictionaryMaxLength);
                                    dictionaryWrite %= CircularDictionaryMaxLength; //2.1.3.1.4
                                    offset++;
                                }
                                else
                                {
                                    //reference bit, create word from two bytes - note big-Endian ordering
                                    var word = (data[i + offset] << 8) | data[i + offset + 1];

                                    //get the offset into the dictionary
                                    var upper = (word & 0xFFF0) >> 4;

                                    //get the length of bytes to copy
                                    var lower = (word & 0xF) + 2;

                                    if (upper > dictionaryEnd)
                                        throw new XstException("Input stream is corrupt: invalid dictionary reference");

                                    if (upper == dictionaryWrite)
                                        //special dictionary reference means that decompression is complete
                                        return destination;

                                    //cannot just copy the bytes over because the dictionary is a
                                    //circular array so it must properly wrap to beginning
                                    for (int k = 0; k < lower; k++)
                                    {
                                        int correctedOffset = (upper + k) % CircularDictionaryMaxLength; //2.1.3.1.4

                                        if (destination.Position == header.rawSize)
                                            //this is the last token, the rest is just padding
                                            return destination;

                                        destination.WriteByte(dictionary[correctedOffset]);
                                        dictionary[dictionaryWrite++] = dictionary[correctedOffset];
                                        if (dictionaryWrite > dictionaryEnd)
                                            dictionaryEnd = Math.Min(dictionaryWrite, CircularDictionaryMaxLength);
                                        dictionaryWrite %= CircularDictionaryMaxLength; //2.1.3.1.4
                                    }
                                    offset += 2;
                                }
                            }
                            //run is processed
                            i += offset;
                        }
                    }
                    catch (IndexOutOfRangeException)
                    {
                        throw new XstException("Input stream is corrupt: index out of range");
                    }

                    destination.Dispose();
                    break;

                default:
                    throw new XstException("Input stream is corrupt: unknown compression type");
            }

            return null;
        }

        private UInt32 CalculateCrc(byte[] buffer, int offset)
        {
            //2.1.3.2
            return Crc32.Compute(buffer, offset, buffer.Length - offset);
        }

        //enumerations
        private enum CompressionTypes : UInt32
        {
            //2.1.3.1.1
            Compressed = 0x75465A4C,
            UnCompressed = 0x414C454D,
        }

        //nested types
        [StructLayout(LayoutKind.Sequential, Pack = 1)]
        private struct RtfHeader
        {
            //fields
            public UInt32 compSize;
            public UInt32 rawSize;
            public UInt32 compType;
            public UInt32 crc;
        }
    }
}
