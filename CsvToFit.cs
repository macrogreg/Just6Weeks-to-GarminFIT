/// <summary>
/// Authored by MACROGREG. See: https://github.com/macrogreg/Just6Weeks-to-GarminFIT
/// Convert a CSV table with gym workout data to a FIT file in order to import into Garmin Connect.
/// </summary>

using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using Dynastream.Fit;
using FIT = Dynastream.Fit;

namespace Just6Weeks_to_GarminFIT;

public static class ProductInfo
{
    public const ushort Manufacturer = FIT.Manufacturer.Development;
    public const ushort ProductCode = 12345;

    // Max 20 Chars (?)
    //                                0         1         2         3
    //                                0123456789012345678901234567890
    public const string ProductName = "Just6Weeks-to-FIT";
    public const int SerialNumber = 1234567890;
    public const float SoftwareVersion = 0.1f;
    public static readonly Guid DeveloperId = new("c8130744-a919-45f2-a87e-cf5f224ec726");
    public const string SportProfileName = "Just6Weeks Workout";
}

internal static class Extensions
{
    public static FIT.DateTime AsFit(this System.DateTime dateTime)
    {
        return new FIT.DateTime(dateTime);
    }

    public static string PrettyPrint(this TimeSpan span)
    {
        return (span.TotalHours >= 1) ? span.ToString(@"hh\:mm\:ss") : span.ToString(@"mm\:ss");
    }

    public static TimeSpan ParseTimeSpan(this string s, string parseMoniker = null)
    {
        ArgumentNullException.ThrowIfNull(s);
        parseMoniker = String.IsNullOrWhiteSpace(parseMoniker) ? "String to parse" : parseMoniker;

        string[] parts = s.Split(':');
        return parts.Length switch
        {
            2 => new TimeSpan(0, Int32.Parse(parts[0]), Int32.Parse(parts[1])),
            3 => new TimeSpan(Int32.Parse(parts[0]), Int32.Parse(parts[1]), Int32.Parse(parts[2])),
            _ => throw new Exception($"{parseMoniker} (=\"{s}\") has {parts.Length} parts separated by ':', but 2 or 3 are expected.")
        };
    }
}

internal record Just6WeeksWorkout
{
    private readonly Dictionary<int, int> _setReps = new();
    private readonly Dictionary<int, TimeSpan> _lapTimes = new();

    public System.DateTime StartTime { get; set; }

    public string Type { get; set; } = "unspecified";
    public string TypeNorm => (String.IsNullOrWhiteSpace(Type) ? "unspecified" : Type).ToLowerInvariant();

    public int TrainingWeek { get; set; } = 0;
    public int DayInTrainingWeek { get; set; } = 0;

    public TimeSpan TotalTime { get; set; } = TimeSpan.Zero;

    public int GetSetReps(int setNum) => _setReps.TryGetValue(setNum, out int reps) ? reps : 0;
    public TimeSpan GetLapTime(int lapNum) => _lapTimes.TryGetValue(lapNum, out TimeSpan time) ? time : TimeSpan.Zero;
    public void SetSetReps(int setNum, int reps) => _setReps[setNum] = reps;
    public void SetLapTime(int lapNum, TimeSpan time) => _lapTimes[lapNum] = time;

    public int SumOfSets { get; set; } = 0;
    public TimeSpan SumOfLaps { get; set; } = TimeSpan.Zero;

    public int Kcal { get; set; } = 0;

    public System.DateTime GetEndTime() => StartTime.Add(TotalTime);

    public TimeSpan GetActivePartDuration(int setOrLapNum)
    {
        if (_lapTimes.TryGetValue(setOrLapNum, out TimeSpan time))
        {
            return time;
        }

        // Assume, 1 push-up/sit-up takes 1.1 secs. This is based on actual numbers for a many-push-ups set.
        // Likely for sit-ups, this is an overestimation.
        return (1 <= setOrLapNum && setOrLapNum <= 5 && !TypeNorm.Equals("plank"))
            ? TimeSpan.FromSeconds(GetSetReps(setOrLapNum) * 1.1)
            : TimeSpan.Zero;        
    }

    public TimeSpan GetTotalActiveDuration()
    {
        if (SumOfLaps > TimeSpan.Zero)
        {
            return SumOfLaps;
        }

        TimeSpan total = TimeSpan.Zero;
        for (int part = 1; part <= 5; part++)
        {
            total = total + GetActivePartDuration(part);
        }
        return total;
    }


    public TimeSpan GetBreakDuration()
    {
        TimeSpan totalActive = GetTotalActiveDuration();
        return (TotalTime <= totalActive)
            ? TimeSpan.Zero
            : (TotalTime - totalActive) / 4;
    }
    

    public override string ToString()
    {
        return $"Start='{StartTime:yy-MM-dd HH:mm:ss}'; Type='{Type}'; Week={TrainingWeek}; Day={DayInTrainingWeek};"
             + $" Time='{TotalTime.PrettyPrint()}';"
             + $" Sets={{{PrintSet(1)}, {PrintSet(2)}, {PrintSet(3)}, {PrintSet(4)}, {PrintSet(5)}}};"
             + $" SumOfSets={PrintSumOfSets()}; Kcal={Kcal}; Breaks='{GetBreakDuration().PrettyPrint()}'";

    }
    
    public string PrintSet(int setNum) => TypeNorm.Equals("plank") ? GetLapTime(setNum).PrettyPrint() : $"{GetSetReps(setNum)}";
    public string PrintSumOfSets() => TypeNorm.Equals("plank") ? SumOfLaps.PrettyPrint() : $"{SumOfSets}";
}

internal record Just6WeeksSession
{
    private readonly List<Just6WeeksWorkout> _workouts = new();
    public IReadOnlyList<Just6WeeksWorkout> Workouts => _workouts;
    public System.DateTime StartTime { get; private set; }
    public System.DateTime EndTime { get; private set; }

    public void AddWorkout(Just6WeeksWorkout workout)
    {
        ArgumentNullException.ThrowIfNull(workout);

        _workouts.Add(workout);
        
        StartTime = (_workouts.Count > 1 && StartTime <= workout.StartTime) ? StartTime : workout.StartTime;
        EndTime = (_workouts.Count > 1 && EndTime >= workout.GetEndTime()) ? EndTime : workout.GetEndTime();
    }
}

public class CsvToFit
{
    private int _countReadErrors = 0;
    private int _countWriteErrors = 0;
    private int _countWorkoutsSkipped = 0;
    private int _countSessionsSkipped = 0;
  
    static void Main(string[] args)
    {
        (new CsvToFit()).Run(args);
    }    

    public void Run(string[] args)
    {
        const string AppName = "CsvToFit";
        Console.WriteLine();
        Console.WriteLine(AppName);

        DateTimeOffset appStartTime = DateTimeOffset.Now;

        _countReadErrors = _countWriteErrors = _countWorkoutsSkipped = _countSessionsSkipped = 0;

        ArgumentNullException.ThrowIfNull(args);

        if (args.Length != 1)
        {
            Console.WriteLine($"    Usage: {AppName} InputFileName");
            Console.WriteLine($"    Expected 1 argument, received {args.Length}. Exiting.");
            return;
        }

        string inFileName = args[0];
        Console.WriteLine($"Reading '{inFileName}'...");

        IReadOnlyList<Just6WeeksSession> sessions;
        FileStream inFile = new(inFileName, FileMode.Open, FileAccess.Read, FileShare.Read);
        using (inFile)
        {
            sessions = ReadInput(inFile);
        }

        string outFileDir = GetOutFileDir(inFileName);

        Console.WriteLine();
        Directory.CreateDirectory(outFileDir);
        Console.WriteLine($"Writing results to '{outFileDir}'...");

        int countWrittenWorkouts = 0;
        int countWrittenSessions = 0;
        for (int s = 0; s < sessions.Count; s++)
        {
            Just6WeeksSession session = sessions[s];
            string outFileName = $"Session_" + session.EndTime.ToString("yyyy-MM-dd_HH-mm")
                               + $"_w{session.Workouts[0].TrainingWeek:D2}d{session.Workouts[0].DayInTrainingWeek:D2}"
                               + ".fit";
            string outFilePath = Path.Combine(outFileDir, outFileName);

            // if (!outFileName.Equals("Session_w02d02_2022-11-16_21-38.fit"))
            // {
            //     Console.WriteLine($"Skipping session {s + 1}/{sessions.Count} in '{outFilePath}'."
            //                      + " (Likely a debug set-up; see code.)");
            //     _countSessionsSkipped++;
            //     continue;
            // }

            Console.WriteLine($"\nExporting session {s + 1}/{sessions.Count} to '{outFilePath}'.");

            FileStream outFile = new(outFilePath, FileMode.Create, FileAccess.ReadWrite, FileShare.Read);
            using (outFile)
            {
                WriteActivityFile(outFile, session, appStartTime.AddSeconds(countWrittenSessions), out int writtenSessionWorkouts);
                countWrittenWorkouts += writtenSessionWorkouts;
            }

            countWrittenSessions++;
            Console.WriteLine($"Finished writing '{outFileName}'.");
        }

        Console.WriteLine($"\nFinished exporting {countWrittenWorkouts} workouts in {countWrittenSessions} session groups to '{outFileDir}'.");

        if (_countWorkoutsSkipped > 0 || _countSessionsSkipped > 0)
        {
            Console.WriteLine();
            Console.WriteLine(" ~~  ~ Some data was skipped (summary below).      ~ ~~");
            Console.WriteLine(" ~~  ~ See console log above for details.          ~ ~~");
        }        

        if (_countWorkoutsSkipped > 0)
        {
            Console.WriteLine($"\nWorkouts parsed, but skipped: {_countWorkoutsSkipped}.");
        }

        if (_countSessionsSkipped > 0)
        {
            Console.WriteLine($"\nSession groups read, but skipped: {_countSessionsSkipped}.");
        }

        if (_countReadErrors > 0 || _countWriteErrors > 0)
        {
            Console.WriteLine();
            Console.WriteLine(" !!  * Processing Errors detected (summary below). * !!");
            Console.WriteLine(" !!  * Results may be not fully valid.             * !!");
            Console.WriteLine(" !!  * See console log above for details.          * !!");
        }

        if (_countReadErrors > 0)
        {
            Console.WriteLine($"\nErrors while reading and parsing input data: {_countReadErrors}.");
        }
        
        if (_countWriteErrors > 0) {
            Console.WriteLine($"\nErrors while converting to FIT or writing output files: {_countWriteErrors}.");
        }
    }

    private static string GetOutFileDir(string inFilePath)
    {        
        // Normalize and extract parts:
        string directory = Path.GetDirectoryName(inFilePath) ?? "";
        string baseName = Path.GetFileNameWithoutExtension(inFilePath);
        string baseDir = Path.Combine(directory, baseName);

        // If dir already exists, add suffix:
        if (Directory.Exists(baseDir))
        {
            int index = 1;
            string candidate = $"{baseDir} ({index})";
            while (Directory.Exists(candidate))
            {
                candidate = $"{baseDir} ({++index})";
            }
            baseDir = candidate;
        }

        // Dir determined:        
        return baseDir;
    }

    private IReadOnlyList<Just6WeeksSession> ReadInput(Stream inStr)
    {
        const string CsvSeparatorChar = ';';
        Console.WriteLine();

        // Check for valid In Stream:
        ArgumentNullException.ThrowIfNull(inStr);
        if (!inStr.CanRead)
        {
            throw new ArgumentException($"${nameof(inStr)} must support Reads, but it does not.");
        }

        if (!inStr.CanSeek)
        {
            throw new ArgumentException($"${nameof(inStr)} must support Seeking, but it does not.");
        }

        List<Just6WeeksWorkout> workouts = new();

        using StreamReader reader = new(inStr);

        int lineNum = 0;
        string line = reader.ReadLine();

        Console.WriteLine("Skipping first file line: It should contain columns headers. Check:");
        Console.WriteLine(line ?? " - NULL -");

        line = reader.ReadLine();
        while (line != null)
        {
            lineNum++;

            try
            {
                string[] lineFields = line.Split(CsvSeparatorChar);
                if (lineFields.Length < 14)
                {
                    throw new Exception($"Line {lineNum} has {lineFields.Length} columns, but at least 14 were expected.");
                }

                Just6WeeksWorkout workout = new();

                try
                {
                    workout.StartTime = System.DateTime.ParseExact(
                        lineFields[0],
                        "M/d/yyyy h:mm tt",
                        CultureInfo.InvariantCulture
                    );
                }
                catch (Exception ex)
                {
                    throw new Exception($"Cannot parse StartTime ({lineFields[0]})", ex);
                }

                workout.Type = lineFields[1].Trim();

                // Skip [2] (Goal) and [3] (period)

                try
                {
                    workout.TrainingWeek = Int32.Parse(lineFields[4]);
                }
                catch (Exception ex)
                {
                    throw new Exception($"Cannot parse TrainingWeek ({lineFields[4]})", ex);
                }

                try
                {
                    workout.DayInTrainingWeek = Int32.Parse(lineFields[5]);
                }
                catch (Exception ex)
                {
                    throw new Exception($"Cannot parse DayInTrainingWeek ({lineFields[5]})", ex);
                }

                try
                {                    
                    workout.TotalTime = lineFields[6].ParseTimeSpan("TotalTime");
                    //Console.WriteLine($"Line {lineNum}: TotalTime='{workout.TotalTime}'; Raw='{lineFields[6]}'; parts.Len={totTmParts.Length}");
                }
                catch (Exception ex)
                {
                    throw new Exception($"Cannot parse TotalTime ({lineFields[6]})", ex);
                }

                for (int set = 1; set <= 5; set++)
                {
                    try
                    {
                        switch (workout.TypeNorm)
                        {
                            case "plank":
                                workout.SetLapTime(set, lineFields[6 + set].ParseTimeSpan($"Set{set}")); break;
                            default:
                                workout.SetSetReps(set, Int32.Parse(lineFields[6 + set])); break;
                        }
                    }
                    catch (Exception ex)
                    {
                        throw new Exception($"Cannot parse Set{set} ({lineFields[6 + set]})", ex);
                    }
                }

                try
                {
                    switch (workout.TypeNorm)
                    {
                        case "plank":
                            workout.SumOfLaps = lineFields[12].ParseTimeSpan("SumOfSets"); break;
                        default:
                            workout.SumOfSets = Int32.Parse(lineFields[12]); break;
                    }
                }
                catch (Exception ex)
                {
                    throw new Exception($"Cannot parse SumOfSets ({lineFields[12]})", ex);
                }

                try
                {
                    workout.Kcal = Int32.Parse(lineFields[13]);
                }
                catch (Exception ex)
                {
                    throw new Exception($"Cannot parse Kcal ({lineFields[13]})", ex);
                }

                workouts.Add(workout);
            }
            catch (Exception ex)
            {
                _countReadErrors++;
                Console.WriteLine();
                Console.WriteLine($"{ex.GetType().Name} while processing line {lineNum}:");
                Console.WriteLine(ex.ToString());
                Console.WriteLine($"Proceeding to input file next line.");
            }

            line = reader.ReadLine();
        }  // while (line != null)

        if (workouts.Count == 0)
        {
            Console.WriteLine($"No workouts read.");
            return new List<Just6WeeksSession>();
        }

        Console.WriteLine($"{workouts.Count} workouts read. Grouping...");

        // Sort before grouping in sessions:
        workouts.Sort((w1, w2) => w1.StartTime.CompareTo(w2.StartTime));

        // Group workouts into sessions:

        TimeSpan maxGroupStartGap = TimeSpan.FromHours(2);

        List<Just6WeeksSession> sessions = new();

        sessions.Add(new Just6WeeksSession());
        sessions[^1].AddWorkout(workouts[0]);
        // Console.WriteLine($"Workout {1} added to initial session group:");
        // Console.WriteLine("  " + workouts[0].ToString());

        for (int w = 1; w < workouts.Count; w++)
        {
            System.DateTime prevStart = workouts[w - 1].StartTime;
            System.DateTime currStart = workouts[w].StartTime;
            if ((currStart - prevStart) <= maxGroupStartGap
                    && workouts[w].TrainingWeek == workouts[w - 1].TrainingWeek
                    && workouts[w].DayInTrainingWeek == workouts[w - 1].DayInTrainingWeek)
            {
                sessions[^1].AddWorkout(workouts[w]);
                // Console.WriteLine($"Workout {w + 1} added to group #{sessions.Count} as member #{sessions[^1].Workouts.Count}:");
                // Console.WriteLine("  " + workouts[w].ToString());
            }
            else
            {
                sessions.Add(new Just6WeeksSession());
                sessions[^1].AddWorkout(workouts[w]);
                // Console.WriteLine($"Workout {w + 1} added to new group #{sessions.Count}.");
                // Console.WriteLine("  " + workouts[w].ToString());
            }
        }

        Console.WriteLine($"Done reading input. {workouts.Count} workouts in {sessions.Count} session groups.");
        return sessions;
    }

    internal static int ParseCountOrTimespan(string rawStr, string fieldName, string workoutTypeNorm)
    // (Not currently used but kept for future referecne)
    {        
        try
        {
            if (workoutTypeNorm.Equals("plank"))
            {
                string[] tmParts = rawStr.Split(':');
                return (int)Math.Round((tmParts.Length switch
                {
                    2 => new TimeSpan(0, Int32.Parse(tmParts[0]), Int32.Parse(tmParts[1])),
                    3 => new TimeSpan(Int32.Parse(tmParts[0]), Int32.Parse(tmParts[1]), Int32.Parse(tmParts[1])),
                    _ => throw new Exception($"{fieldName} for 'plank' is a TimeSpan;"
                            + $" {fieldName.Length - 1} separators (':') found, but 2 or 3 are expected.")
                }).TotalSeconds);
            }
            else
            {
                return Int32.Parse(rawStr);
            }
        }
        catch (Exception ex)
        {
            throw new Exception($"Cannot parse {fieldName} ({rawStr})", ex);
        }
    }

    private void WriteActivityFile(
        Stream outStr,
        Just6WeeksSession session,
        DateTimeOffset sessionCreationTimetamp,
        out int countWrittenWorkouts
    )
    {
        // The `session` contains one or more workouts.
        // Each `workout` contains 5 sets.
        // Each set either has reps (e.g. push-up), or it has a duration (plank).
        //
        // We create one FIT file per `session`. The file consists of blocks called messages ordered as follows:
        //
        //       (headers:)
        //     FileIdMesg:          Required file metadata.
        //     DeviceInfoMesg:      Additional file metadata (optional, but we include it).
        //
        //       (custom field metadata:)
        //     DeveloperDataIdMesg: Metadata about App Developer (required for custom fields)
        //     For each custom field (we have 2, TrainingWeek and DayInTrainingWeek):
        //         FieldDescriptionMesg: Metadata about the custom field type.
        //     
        //       (workout sets data:)
        //     For each workout in the session, we write 5+4=9 messages:                
        //         SetMesg: Activity kind (push-up, sit-up, plank, ...), number of reps, etc.
        //         SetMesg: Describe the rest break after the prev set (skipped after last set).
        //
        //       (summary data:)
        //     SessionMesg: The summary of the entire session, including
        //         - total times stats,
        //         - sport kind (FitnessEquipment/StrengthTraining)
        //         - display title
        //     ActivityMesg: Required summary message (it exists because of multi-sport FITs whcih have multiple sessions).

        Console.WriteLine($"Session '{session.StartTime}'..'{session.EndTime}' with {session.Workouts.Count} workouts:" );

        // Check for valid Out Stream:
        
        ArgumentNullException.ThrowIfNull(outStr);
        if (!outStr.CanRead)
        {
            throw new ArgumentException($"${nameof(outStr)} must support Reads, but it does not.");
        }

        if (!outStr.CanSeek)
        {
            throw new ArgumentException($"${nameof(outStr)} must support Seeking, but it does not.");
        }

        if (!outStr.CanWrite)
        {
            throw new ArgumentException($"${nameof(outStr)} must support Writes, but it does not.");
        }

        // Create a FIT Encode object:

        Encode encoder = new(ProtocolVersion.V20);
        encoder.Open(outStr);

        try
        {
            // Write DeviceId and DeviceInfo messages (headers):

            WriteHeaderMessages(encoder, sessionCreationTimetamp);

            // Write messages DeveloperDataId and FieldDescription messages (custom field metadata):

            DeveloperDataIdMesg developerDataIdMesg = WriteDevDataIdMessage(encoder);

            (
                FieldDescriptionMesg trainWeekFieldDescMesg,
                FieldDescriptionMesg trainDayInWeekFieldDescMesg
            ) = WriteFieldDescriptorMessages(
                encoder,
                (byte)developerDataIdMesg.GetDeveloperDataIndex()
            );

            // Work each workout we will write several Set of Lap messages (workout sets data):

            int sessionSetNum = 0;
            countWrittenWorkouts = 0;
            for (int wN = 0; wN < session.Workouts.Count; wN++)
            {
                Just6WeeksWorkout workout = session.Workouts[wN];
                try
                {
                    Console.WriteLine($"  [{wN + 1}/{session.Workouts.Count}] {workout}");

                    if (!TryGetFitExerciseCategory(workout, out ushort exerciseCategory, out ushort exerciseNameCode))
                    {
                        Console.WriteLine($"    Unsupported workout type (normalized: \"{workout.TypeNorm}\".");
                        Console.WriteLine($"    Skipping workout. Will process other workouts in the current session, if any.");
                        _countWorkoutsSkipped++;
                        continue;
                    }

                    System.DateTime setStartTime = workout.StartTime;
                    for (int set = 1; set <= 5; set++)
                    {
                        // Write Set data:
                        SetMesg setMesg = new();
                        setMesg.SetMessageIndex((ushort)(sessionSetNum++));

                        setMesg.SetSetType(SetType.Active);
                        setMesg.SetCategory(0, exerciseCategory);
                        setMesg.SetCategorySubtype(0, exerciseNameCode);

                        if (workout.TypeNorm is "plank")
                        {
                            // Garmin connect aggregates set-based exercise reports only by reps (or by weight), but not by time.
                            // So we pretend seconds are reps, so that we can track progress using reports.
                            setMesg.SetRepetitions((ushort)workout.GetLapTime(set).TotalSeconds);
                        }
                        else
                        {
                            setMesg.SetRepetitions((ushort)workout.GetSetReps(set));
                        }

                        System.DateTime setEndTime = setStartTime.Add(workout.GetActivePartDuration(set));
                        setMesg.SetStartTime(setStartTime.AsFit());
                        setMesg.SetTimestamp(setEndTime.AsFit());
                        setMesg.SetDuration(setMesg.GetTimestamp().GetTimeStamp() - setMesg.GetStartTime().GetTimeStamp());

                        encoder.Write(setMesg);
                        setStartTime = setEndTime;

                        // Write Rest/Break data:
                        if (set < 5)
                        {
                            setMesg = new();
                            setMesg.SetMessageIndex((ushort)(sessionSetNum++));

                            setMesg.SetSetType(SetType.Rest);

                            setEndTime = setStartTime.Add(workout.GetBreakDuration());
                            setMesg.SetStartTime(setStartTime.AsFit());
                            setMesg.SetTimestamp(setEndTime.AsFit());
                            setMesg.SetDuration(setMesg.GetTimestamp().GetTimeStamp() - setMesg.GetStartTime().GetTimeStamp());

                            encoder.Write(setMesg);
                            setStartTime = setEndTime;
                        }
                    }

                    countWrittenWorkouts++;
                }
                catch (Exception ex)
                {
                    _countWriteErrors++;
                    Console.WriteLine();
                    Console.WriteLine($"  {ex.GetType().Name} while writing session workout ({wN + 1}/{session.Workouts.Count}):");
                    Console.WriteLine("  " + ex.ToString());
                    Console.WriteLine("  Proceeding to the next workout, but:");
                    Console.WriteLine("  The *entire* session .FIT file may be corrupt. Fix the issue and regenerate it.");
                }
            }

            // Every FIT ACTIVITY file MUST contain at least one Session message

            TimeSpan totalElapsedTime = session.Workouts
                .Where( w => TryGetFitExerciseCategory(w, out _, out _))
                .Select(w => w.TotalTime)
                .Aggregate(TimeSpan.Zero, (a, b) => a + b);

            TimeSpan totalActiveTime = session.Workouts
                .Where( w => TryGetFitExerciseCategory(w, out _, out _))
                .Select(w => w.GetTotalActiveDuration())
                .Aggregate(TimeSpan.Zero, (a, b) => a + b);            

            Console.WriteLine($"  Summary:"
                    + $" Total Sets: {sessionSetNum};"
                    + $" '{session.EndTime}' - '{session.StartTime}' = '{session.EndTime - session.StartTime}';"
                    + $" Total Elapsed Agg: '{totalElapsedTime}';"
                    + $" Active Agg: '{totalActiveTime}';");

            int trainingWeek = session.Workouts[0].TrainingWeek;
            int dayInTrainingWeek = session.Workouts[0].DayInTrainingWeek;

            SessionMesg sessionMesg = new();
            sessionMesg.SetMessageIndex(0);
            sessionMesg.SetSportProfileName($"{ProductInfo.SportProfileName} (week {trainingWeek} / day {dayInTrainingWeek})");

            sessionMesg.SetStartTime(session.StartTime.AsFit());
            sessionMesg.SetTimestamp(session.EndTime.AsFit());

            sessionMesg.SetTotalElapsedTime((float) totalElapsedTime.TotalSeconds);
            sessionMesg.SetTotalTimerTime((float) totalActiveTime.TotalSeconds);            

            sessionMesg.SetSport(Sport.FitnessEquipment);
            sessionMesg.SetSubSport(SubSport.StrengthTraining);

            sessionMesg.SetFirstLapIndex(0);
            sessionMesg.SetNumLaps((ushort) sessionSetNum);

            // Set training Week and Day via developer fields
            DeveloperField trainWeekField = new(trainWeekFieldDescMesg, developerDataIdMesg);
            trainWeekField.SetValue(trainingWeek);
            sessionMesg.SetDeveloperField(trainWeekField);

            DeveloperField trainDayInWeekField = new(trainDayInWeekFieldDescMesg, developerDataIdMesg);
            trainDayInWeekField.SetValue(dayInTrainingWeek);
            sessionMesg.SetDeveloperField(trainDayInWeekField);

            encoder.Write(sessionMesg);

            // Every FIT ACTIVITY file MUST contain EXACTLY one Activity message
            ActivityMesg activityMesg = new();
            activityMesg.SetNumSessions(1);

            activityMesg.SetTimestamp(session.EndTime.AsFit());
            activityMesg.SetLocalTimestamp(session.EndTime.AsFit().GetTimeStamp());
            activityMesg.SetTotalTimerTime((float) totalActiveTime.TotalSeconds);        

            encoder.Write(activityMesg);
        }
        finally
        {
            // Update the data size in the header and calculate the CRC
            encoder.Close();
        }
    }

    private static bool TryGetFitExerciseCategory(Just6WeeksWorkout workout, out ushort exerciseCategory, out ushort exerciseNameCode)
    {
        switch (workout.TypeNorm)
        {
            case "plank":
                exerciseCategory = ExerciseCategory.Plank;
                exerciseNameCode = PlankExerciseName.Plank;
                return true;

            case "pull-ups":
                exerciseCategory = ExerciseCategory.PullUp;
                exerciseNameCode = PullUpExerciseName.PullUp;
                return true;

            case "push-ups":
                exerciseCategory = ExerciseCategory.PushUp;
                exerciseNameCode = PushUpExerciseName.PushUp;
                return true;

            case "sit-ups":
                exerciseCategory = ExerciseCategory.SitUp;
                exerciseNameCode = SitUpExerciseName.SitUp;
                return true;

            case "squats":
                exerciseCategory = ExerciseCategory.Squat;
                exerciseNameCode = SquatExerciseName.Squat;
                return true;

            default:
                exerciseCategory = ExerciseCategory.Invalid;
                exerciseNameCode = (ushort)0xFFFF;
                return false;
        }
    }

    private static void WriteHeaderMessages(Encode encoder, DateTimeOffset sessionCreationTimetamp)
    {
        // Every FIT file MUST contain a File ID message
        FileIdMesg fileIdMesg = new();
        fileIdMesg.SetType(FIT.File.Activity);
        fileIdMesg.SetManufacturer(ProductInfo.Manufacturer);
        fileIdMesg.SetProduct(ProductInfo.ProductCode);
        fileIdMesg.SetProductName(ProductInfo.ProductName);
        fileIdMesg.SetSerialNumber(ProductInfo.SerialNumber);
        fileIdMesg.SetTimeCreated(sessionCreationTimetamp.DateTime.AsFit());

        encoder.Write(fileIdMesg);

        // A Device Info message is a BEST PRACTICE for FIT ACTIVITY files
        DeviceInfoMesg deviceInfoMesg = new();
        deviceInfoMesg.SetDeviceIndex(DeviceIndex.Creator);
        deviceInfoMesg.SetManufacturer(ProductInfo.Manufacturer);
        deviceInfoMesg.SetProduct(ProductInfo.ProductCode);
        deviceInfoMesg.SetProductName(ProductInfo.ProductName);
        deviceInfoMesg.SetSerialNumber(ProductInfo.SerialNumber);
        deviceInfoMesg.SetSoftwareVersion(ProductInfo.SoftwareVersion);
        deviceInfoMesg.SetTimestamp(sessionCreationTimetamp.DateTime.AsFit());

        encoder.Write(deviceInfoMesg);
    }

    private static DeveloperDataIdMesg WriteDevDataIdMessage(Encode encoder)
    {
        // Create the Developer Id message for the developer data fields.
        DeveloperDataIdMesg devIdMesg = new();

        byte[] appId = ProductInfo.DeveloperId.ToByteArray();
        for (int offs = 0; offs < appId.Length; offs++)
        {
            devIdMesg.SetApplicationId(offs, appId[offs]);
        }

        devIdMesg.SetDeveloperDataIndex(0);
        devIdMesg.SetApplicationVersion(10); // = 0.1

        encoder.Write(devIdMesg);
        return devIdMesg;
    }

    private static
    (
        FieldDescriptionMesg trainWeekFieldDescMesg,
        FieldDescriptionMesg trainDayInWeekFieldDescMesg
    ) WriteFieldDescriptorMessages(Encode encoder, byte devDataIndex)
    {
        // Create the Developer Data Field Descriptions

        FieldDescriptionMesg trainingWeekNumDescMesg = new();
        trainingWeekNumDescMesg.SetDeveloperDataIndex(devDataIndex);
        trainingWeekNumDescMesg.SetFieldDefinitionNumber(0);
        trainingWeekNumDescMesg.SetFitBaseTypeId(FitBaseType.Uint16);
        trainingWeekNumDescMesg.SetFieldName(0, "Week Num");
        trainingWeekNumDescMesg.SetUnits(0, "weeks");
        trainingWeekNumDescMesg.SetNativeMesgNum(MesgNum.Session);

        FieldDescriptionMesg trainingDayInWeekNumDescMesg = new();
        trainingDayInWeekNumDescMesg.SetDeveloperDataIndex(devDataIndex);
        trainingDayInWeekNumDescMesg.SetFieldDefinitionNumber(1);
        trainingDayInWeekNumDescMesg.SetFitBaseTypeId(FitBaseType.Uint8);
        trainingDayInWeekNumDescMesg.SetFieldName(0, "Day in Week");
        trainingDayInWeekNumDescMesg.SetUnits(0, "days");
        trainingDayInWeekNumDescMesg.SetNativeMesgNum(MesgNum.Session);

        encoder.Write(trainingWeekNumDescMesg);
        encoder.Write(trainingDayInWeekNumDescMesg);

        return (trainingWeekNumDescMesg, trainingDayInWeekNumDescMesg);
    }
}
