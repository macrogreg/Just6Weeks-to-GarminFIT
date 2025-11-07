using System;
using System.IO;
using System.Globalization;
using Dynastream.Fit;

public class Version1
{
    public static void Run(string[] args)
    {
        const string workDir = @"Just6Weeks-to-GarminFIT\";
        string inputFile = workDir + "workouts.csv";
        string outputFile = workDir + "workouts.fit";

        using (FileStream fitDest = new FileStream(outputFile, FileMode.Create, FileAccess.ReadWrite, FileShare.None))
        {
            Encode encoder = new Encode(ProtocolVersion.V20);
            encoder.Open(fitDest);

            // File ID message
            FileIdMesg fileId = new FileIdMesg();
            fileId.SetType(Dynastream.Fit.File.Activity);
            fileId.SetManufacturer(Manufacturer.Development);
            fileId.SetProduct(1);
            fileId.SetTimeCreated(new Dynastream.Fit.DateTime(System.DateTime.UtcNow));
            encoder.Write(fileId);

            System.DateTime sysEnd = System.DateTime.Now;

            using (var reader = new StreamReader(inputFile))
            {
                string header = reader.ReadLine(); // skip header
                string line;

                int remainingLines = 1;
                while (remainingLines > 0 && (line = reader.ReadLine()) != null)
                {
                    remainingLines--;

                    string[] cols = line.Split(';');
                    if (cols.Length < 14) continue;

                    // Parse timestamp from "Date" column
                    System.DateTime sysStart = System.DateTime.ParseExact(
                        cols[0],
                        "M/d/yyyy h:mm tt",
                        CultureInfo.InvariantCulture
                    );

                    sysStart = new System.DateTime(2025, 10, 3, 17, 00, 00);
                    sysEnd = sysStart.AddHours(1);

                    string exercise = cols[1].Trim().ToLower();
                    int kcal = int.TryParse(cols[13], out var k) ? k : 0;

                    // Create session
                    SessionMesg session = new SessionMesg();

                    session.SetStartTime(new Dynastream.Fit.DateTime(sysStart));
                    session.SetTimestamp(new Dynastream.Fit.DateTime(sysEnd));
                    session.SetTotalElapsedTime((float)(sysEnd - sysStart).TotalSeconds);
                    session.SetTotalTimerTime((float)(sysEnd - sysStart).TotalSeconds);

                    session.SetSport((Sport) 17); // ✅ fixed
                    session.SetTotalCalories((ushort)kcal);
                    encoder.Write(session);

                    // Map CSV exercise to Garmin ExerciseCategory + ExerciseName
                    ushort exCat = 0;
                    ushort exName = 0;

                    if (exercise.Contains("push-up"))
                    {
                        exCat = (ushort)ExerciseCategory.PushUp;
                        exName = (ushort)PushUpExerciseName.PushUp;
                    }
                    else if (exercise.Contains("pull-up"))
                    {
                        exCat = (ushort)ExerciseCategory.PullUp;
                        exName = (ushort)PullUpExerciseName.PullUp;
                    }
                    else if (exercise.Contains("sit-up"))
                    {
                        exCat = (ushort)ExerciseCategory.Crunch;
                        // fallback if SitUp missing
                        exName = (ushort)CrunchExerciseName.Crunch;
                    }
                    else if (exercise.Contains("squat"))
                    {
                        exCat = (ushort)ExerciseCategory.Squat;
                        exName = (ushort)SquatExerciseName.Squat;
                    }
                    else if (exercise.Contains("plank"))
                    {
                        exCat = (ushort)ExerciseCategory.Plank;
                        exName = (ushort)PlankExerciseName.Plank;
                    }
                    else if (exercise.Contains("hand press"))
                    {
                        exCat = (ushort)ExerciseCategory.Unknown;
                        exName = 0;
                    }
                    else
                    {
                        exCat = (ushort)ExerciseCategory.Unknown;
                        exName = 0;
                    }

                    // Write an ExerciseTitle message
                    ExerciseTitleMesg exTitle = new ExerciseTitleMesg();
                    exTitle.SetExerciseCategory(exCat);
                    exTitle.SetExerciseName(exName);
                    encoder.Write(exTitle);

                    // Now add sets
                    for (int i = 7; i <= 11; i++) // Set 1..5 columns
                    {
                        string raw = cols[i].Trim();
                        if (string.IsNullOrEmpty(raw)) continue;

                        SetMesg set = new SetMesg();

                        var setStart = sysStart.AddSeconds((i - 6) * 30);
                        set.SetStartTime(new Dynastream.Fit.DateTime(setStart));

                        //var setEnd = sysStart.AddSeconds((i - 6) * 30 + 1);
                        //set.SetTimestamp(new Dynastream.Fit.DateTime(setEnd));
                        //set.SetDuration((float) (setEnd - setStart).TotalSeconds);

                        set.SetCategory(0, exCat); // ✅ use SetCategory
                        set.SetCategorySubtype(0, exName); // ✅ index 0 for subtype list

                        if (exercise.Contains("plank"))
                        {
                            if (TimeSpan.TryParseExact(raw, "m\\:ss", CultureInfo.InvariantCulture, out var ts))
                            {
                                set.SetDuration((float)ts.TotalSeconds);
                            }
                        }
                        else
                        {
                            if (int.TryParse(raw, out int reps))
                            {
                                set.SetRepetitions((ushort)reps);
                            }
                        }

                        encoder.Write(set);
                    }
                }
            }

            // activity (must come after all sessions)
            ActivityMesg activity = new ActivityMesg();
            activity.SetTimestamp(new Dynastream.Fit.DateTime(sysEnd));
            activity.SetType(Activity.Manual);            
            activity.SetNumSessions(1);
            encoder.Write(activity);

            encoder.Close();
        }

        Console.WriteLine($"FIT file created: {outputFile}");
    }
}
