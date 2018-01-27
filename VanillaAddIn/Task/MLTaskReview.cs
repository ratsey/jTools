using System;
using Newtonsoft.Json;
using Newtonsoft.Json.Converters;

namespace JTools.MLO
{
    public class MLTaskReview : MLOTaskProperty
    {
        private ReviewPeriod _reviewEveryPeriod = ReviewPeriod.Weeks;

        public DateTime NextReviewDate { get; set; }
        public bool HasNextReviewDate { get; set; }
        public int ReviewEveryDuration { get; set; }

        [JsonConverter(typeof(StringEnumConverter))]
        public ReviewPeriod ReviewEveryPeriod { get => _reviewEveryPeriod; set => _reviewEveryPeriod = value; }
    }
}