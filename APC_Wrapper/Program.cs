using System;
using System.Diagnostics;       //subprocesses
using System.IO;                //file operations
using System.IO.Compression;    //zip operations
using System.ServiceProcess;    //control services
using System.Threading;         //sleep function

/***********************************************************************
* Project           : APC
*
* Program name      : APC_Wrapper.exe
*
* Author            : James Kinnaird
*
* Date created      : 12/3/2018
*
* Purpose           : Silently installs, runs, and uninstalls APC to gather panel data. 
*                     APC is packaged within this exe. 
*                     ASC will be stopped before APC is run to prevent conflicts and will be restarted afterwards.
*                     
* Return codes      : -1    -   Extraction failure
*                     -2    -   ASC could not be stopped
*                     -3    -   ASC could not be restarted
*                     -4    -   APC exceeded the alotted run window
***********************************************************************/

namespace APC_Wrapper
{
    class Program
    {
        /*
         * Performs actions with a designated service, such as start or stop 
        */
        static bool Service_control(string service_name, string action)
        {
            bool success = false;                                               //if stoping service succeeded
            bool service_exists = false;                                        //if service exists
            
            foreach(var serv in ServiceController.GetServices())
            {
                if(serv.ServiceName == service_name) { service_exists = true; break; } //loop through all services to see if one matches our service_name
            }

            if (service_exists)                                                 //if service exists
            {
                ServiceController service = new ServiceController(service_name);//create service object
                TimeSpan timeout = new TimeSpan(0, 0, 15);                      //amount of time to wait for the service to start or stop
                try
                {
                    if (action == "STOP")                                       //if stopping the service
                    {
                        Console.WriteLine("Stopping " + service_name);
                        if (service.Status != ServiceControllerStatus.Stopped)  //if it is not already stopped
                        {
                            service.Stop();                                     //stop it
                            service.WaitForStatus(ServiceControllerStatus.Stopped, timeout);    //wait for it to stop
                            if (service.Status == ServiceControllerStatus.Stopped) { success = true; } //if service is stopped, return success
                        }
                        else success = true;                                    //if service was already stopped, don't bother stopping it and return success
                    }
                    else if (action == "START")                                 //if starting the service
                    {
                        Console.WriteLine("Starting " + service_name);
                        if (service.Status != ServiceControllerStatus.Running)  //if the service is not already running
                        {
                            service.Start();                                    //start it
                            service.WaitForStatus(ServiceControllerStatus.Running, timeout);    //wait for it to start
                            if (service.Status == ServiceControllerStatus.Running) { success = true; } //if service is started, return success
                            else { Console.WriteLine(service.Status); }
                        }
                        else { success = true; }                                //if the service was already running, return success
                    }
                }
                catch (System.ServiceProcess.TimeoutException te) { Console.WriteLine(te.ToString()); success = false; }    //if operation timed out, return failure
                catch (Exception e) { Console.WriteLine(e.ToString()); }
            }
            else { success = true; }                                            //if service does not exist just return success
            return success;
        }

        static void Main(string[] args)
        {
            string temp_zip_path = "D:\\APC.zip";                          //place to temporarily extract APC zip
            string dest_path = "D:\\APC";                                       //directory to extract APC to
            int exit_code = 0;

            //--------------------------------------------------------------------------------------------------
            //                                     Unpack and unzip APC
            //--------------------------------------------------------------------------------------------------
            try
            {
                Console.WriteLine("Unpacking APC");
                File.WriteAllBytes(temp_zip_path, Properties.Resources.APC);    //unpack APC
            }
            catch(Exception e) { Console.WriteLine(e.ToString()); Environment.Exit(-1); }

            if (File.Exists(temp_zip_path))
            {
                if (Directory.Exists(dest_path)) { Console.WriteLine("Removing previous instance of APC"); Directory.Delete(dest_path, true); }

                Console.WriteLine("Installing APC");
                ZipFile.ExtractToDirectory(temp_zip_path, dest_path);           //unzip APC
                File.Delete(temp_zip_path);                                     //delete temporary zip
            }
            else { Environment.Exit(-1); }

            //--------------------------------------------------------------------------------------------------
            //                                     Run APC
            //--------------------------------------------------------------------------------------------------

            if (Service_control("ASC", "STOP"))                                     //Stop asc if it is running
            {
                //run APC
                DateTime started = DateTime.Now;                                    //what time did the subprocess start
                TimeSpan time_spent = DateTime.Now - started;                       //how long has the process been running

                ProcessStartInfo processInfo = new ProcessStartInfo
                {
                    UseShellExecute = false,                                        //execute directly from executable
                    CreateNoWindow = true,                                          //do not show a window
                    FileName = "D:\\APC\\java\\bin\\java.exe",                      //run this file
                    Arguments = "-Xmx1024m -Xms16m -XX:+UseG1GC -cp D:\\APC\\ABNPanelCheck.jar apc.APC D:\\APC\\panel.ini"      //with these arguments
                };
                
                Console.WriteLine("Running APC");
                Process proc = Process.Start(processInfo);                          //run APC process
                while (time_spent < TimeSpan.FromHours(2))                          //while the process has not been running for more than 2 hours
                {
                    if (proc.HasExited)                                              //if the process has exited
                    {
                        exit_code = proc.ExitCode;
                        Console.WriteLine("Process exited with exit code " + exit_code);    //write the exit code
                        break;
                    }

                    time_spent = DateTime.Now - started;                            //update timespan variable

                    if (time_spent.Hours >= 2)                                      //if process has been running for more than 2 hours
                    {
                        Console.WriteLine("Process has exceeded the time limit");
                        proc.Kill();                                                //kill the task
                        exit_code = -4;                                             //set the time_elapsed flag
                        break;                                                      //break the loop
                    }
                    Thread.Sleep(1 * 1000);
                }
            }
            else { exit_code = -2; }

            if(!Service_control("ASC", "START")) { Console.WriteLine("ASC did not start!"); exit_code = -3; }   //start ASC

            //--------------------------------------------------------------------------------------------------
            //                                     Uninstall APC
            //--------------------------------------------------------------------------------------------------

            Console.WriteLine("Uninstalling APC");
            if (Directory.Exists(dest_path)) { Directory.Delete(dest_path, true); } //delete any extraneous files as necessary
            if (File.Exists(temp_zip_path)) { File.Delete(temp_zip_path); }
            Thread.Sleep(5*1000);
            Environment.Exit(exit_code);  //return APCs error code
        }
    }
}
