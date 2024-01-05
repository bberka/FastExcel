using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Xml;
using System.Xml.Linq;

namespace FastExcel
{
  /// <summary>
  ///   Read and update xl/sharedStrings.xml file
  /// </summary>
  public class SharedStrings
  {
    internal SharedStrings(ZipArchive archive) {
      ZipArchive = archive;
      SharedStringsExists = false;

      StringDictionary = new Dictionary<int, string>();

      var sharedStringsXmlExists = ZipArchive.Entries.Any(entry => entry.FullName == "xl/sharedStrings.xml");
      if (!sharedStringsXmlExists) {
        SharedStringsExists = false;
        return;
      }

      using (var stream = ZipArchive.GetEntry("xl/sharedStrings.xml")?.Open()) {
        if (stream == null) return;

        var document = XDocument.Load(stream) ?? throw new Exception("Failed to load sharedStrings.xml");

        SharedStringsExists = true;
        var i = 0;
        var stringList = document.Descendants().Where(d => d.Name.LocalName == "t").Select(e => XmlConvert.DecodeName(e.Value)).ToList();
        /*
         * NOTE: sharedStrings.xml can contain duplicate strings, but the index is always unique
         * so we don't care about duplicated and just add them to the dictionary by index
         */
        foreach (var currentString in stringList) StringDictionary.Add(i++, currentString);
      }
    }

    //A dictionary is a lot faster than a list
    private Dictionary<int, string> StringDictionary { get; }

    private bool SharedStringsExists { get; }
    private ZipArchive ZipArchive { get; }

    /// <summary>
    ///   Is there any pending changes
    /// </summary>
    public bool PendingChanges { get; private set; }

    /// <summary>
    ///   Is in read/write mode
    /// </summary>
    public bool ReadWriteMode { get; set; }

    internal int AddString(string stringValue) {
      var nextIndex = StringDictionary.Count;
      StringDictionary.Add(nextIndex, stringValue);
      if (!ReadWriteMode) return -1;

      PendingChanges = true;
      return nextIndex;
    }

    internal void Write() {
      // Only update if changes were made
      if (!PendingChanges) return;

      StreamWriter streamWriter = null;
      try {
        streamWriter = SharedStringsExists
                         ? new StreamWriter(ZipArchive.GetEntry("xl/sharedStrings.xml")?.Open() ?? throw new Exception("Failed to open xl/sharedStrings.xml"))
                         //This exception will never throw because we check if the file exists before we get the stream
                         : new StreamWriter(ZipArchive.CreateEntry("xl/sharedStrings.xml").Open());

        // TODO instead of saving the headers then writing them back get position where the headers finish then write from there

        /* Note: the count attribute value is wrong, it is the number of times strings are used thoughout the workbook it is different to the unique count
         *       but because this library is about speed and Excel does not seem to care I am not going to fix it because I would need to read the whole workbook
         */

        var textToWrite = string.Format("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
                                        "<sst uniqueCount=\"{0}\" count=\"{0}\" xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">",
                                        StringDictionary.Count);
        streamWriter.Write(textToWrite);

        // Add Rows
        foreach (var stringValue in StringDictionary) streamWriter.Write($"<si><t>{XmlConvert.EncodeName(stringValue.Value)}</t></si>");

        //Add Footers
        streamWriter.Write("</sst>");
        streamWriter.Flush();
      }
      finally {
        streamWriter?.Dispose();
        PendingChanges = false;
      }
    }

    internal string GetString(string position) {
      if (int.TryParse(position, out var pos)) return GetString(pos + 1);

      // TODO: should I throw an error? this is a corrupted excel document
      throw new Exception("Corrupted excel document position: " + position);
      return string.Empty;
    }

    internal string GetString(int position) {
      var valueExist = StringDictionary.TryGetValue(position - 1, out var value);
      return !valueExist
               //TODO SHOULD THROW ? 
               // throw new Exception("String does not exist in shared strings position: " + position);
               ? string.Empty
               : value;
    }
  }
}