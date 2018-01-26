using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace JTools.MLO
{
    public class MLTaskReview : MLOTaskProperty
    {
        public int NextReview
        {
            get => default(int);
            set
            {
            }
        }

        public int hasNextReviewDate
        {
            get => default(int);
            set
            {
            }
        }

        public int ReviewEveryDuration
        {
            get => default(int);
            set
            {
            }
        }

        public ReviewPeriod ReviewEveryPeriod
        {
            get => ReviewPeriod.Weeks;
            set
            {
            }
        }
    }
}