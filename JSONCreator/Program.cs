using JTools.MLO;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace JSONCreator
{
    class Program
    {
        static void Main(string[] args)
        {
            MLOTask task = new MLOTask();            

            task.Title = "New Task Title";
            task.Note = "Formatted Note";

            task.General.Contexts.Add("#ContextTag1");
            task.General.Contexts.Add("#ContextTag2");

            task.General.Importance = 99;
            task.General.Urgency = 101;

            task.General.Folder = false;
            task.General.HideBranchInToDo = false;
            task.General.CompleteSubTasksInOrder = false;

            task.General.GoalPeriod = GoalPeriod.Week;

            task.General.TextTag = "#TextTag";

            string json = JsonConvert.SerializeObject(task);

            Debug.Print(json);
        }
    }
}
