using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace JTools.MLO
{
    public class MLTaskReview : MLOTaskProperty
    {
        private ReviewPeriod _reviewEveryPeriod = ReviewPeriod.Weeks;

        public DateTime NextReviewDate { get; set; }
        public bool HasNextReviewDate { get; set; }
        public int ReviewEveryDuration { get; set; }

        public ReviewPeriod ReviewEveryPeriod { get => _reviewEveryPeriod; set => _reviewEveryPeriod = value; }
    }
}