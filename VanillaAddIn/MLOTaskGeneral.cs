using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace JTools.MLO
{
    public class MLOTaskGeneral : MLOTaskProperty
    {
        private List<string> _contexts = new List<string>();
        private int _importance = 100;
        private int _urgency = 100;
        private GoalPeriod _goalPeriod = GoalPeriod.None;

        public List<string> Contexts
        {
            get { return _contexts; }
            set { _contexts = value; }
        }

        public int Importance { get => _importance; set => _importance = value; }
        public int Urgency { get => _urgency; set => _urgency = value; }
        public GoalPeriod GoalPeriod { get => _goalPeriod; set => _goalPeriod = value; }
        public string TextTag { get; set; }
        public bool Folder { get; set; }
        public bool HideBranchInToDo { get; set; }
        public bool CompleteSubTasksInOrder { get; set; }
    }
}