using System;
using System.IO;
using System.Globalization;
using Dynastream.Fit;
using Just6Weeks_to_GarminFIT;

class CsvToFit
{
    static void Main(string[] args)
    {
        //Version1.Run(args);
        (new Version2()).Run(args);
        //Version3.Run(args);
    }
}
