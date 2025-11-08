# Just6Weeks-to-GarminFIT

This quick-and-dirty demo shows how to convert a simple CSV table with gym workout data to a FIT file in order to be imported into Garmin Connect.

The [FIT protocol](https://developer.garmin.com/fit/protocol/) is the format that Garmin uses to send exercise data from devices (e.g., a watch) to the software that processes it (on the phone or on their backend). It is also the file format for manually importing exercise data into their end-user portal, [Garmin Connect](https://connect.garmin.com).  

In fact, Garmin Connect can also import other data formats, including .GPX (raw GPS tracks) and .TCX (raw GPS data with a some additional health sensor data and activity metadata in XML format, similar to GPX). However, those formats are primarily targeted at variations of running/cycling/swimming. You cannot use them for in-door gym training data.  
There are only 2 ways to import gym workouts into Garmin connect: type by hand in the browser (or the phone app), or via a FIT file.

The FIT protocol is binary and optimized for streaming form devices. There is an [SDK](https://developer.garmin.com/fit/) for working with FIT data, but it is poorly documented. It turned out to be tricky to create FIT files that would actually successfully import into Garmin Connect.  
This project demonstrates how to generate valid FIT data for multiple gym workouts, each consisting of multiple exercises, each having multiple sets.

This demo uses gym workout data from the 'Just 6 Weeks' mobile app, but any other data origin is generally possible.  

> **What to expect:**  
> This repo contains some quick-and-dirty code focussed on creating a file in the correct format. Feel free to use it ([MIT license](./LICENSE)), but be prepared to do some refactoring. Also, take note of [Garmin's license for the SDK](https://developer.garmin.com/fit/download/).

> **Validation:**  
> The data produced by this tool is correctly imported into Garmin Connect as of Nov 2025. A sample data file is included. This is, just like the code, quick-and-dirty. Any contributions, including more rigorous or automated testing are welcome.



