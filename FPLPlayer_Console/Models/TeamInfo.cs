using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FPLPlayer_Console.Models
{
    class TeamInfo
    {
        public List<Team> teams { get; set; }
    }

    public class CurrentEventFixture
    {
        public bool is_home { get; set; }
        public int day { get; set; }
        public int event_day { get; set; }
        public int month { get; set; }
        public int id { get; set; }
        public int opponent { get; set; }
    }

    public class NextEventFixture
    {
        public bool is_home { get; set; }
        public int day { get; set; }
        public int event_day { get; set; }
        public int month { get; set; }
        public int id { get; set; }
        public int opponent { get; set; }
    }

    public class Team
    {
        public int id { get; set; }
        public List<CurrentEventFixture> current_event_fixture { get; set; }
        public List<NextEventFixture> next_event_fixture { get; set; }
        public string name { get; set; }
        public int code { get; set; }
        public string short_name { get; set; }
        public bool unavailable { get; set; }
        public int strength { get; set; }
        public int position { get; set; }
        public int played { get; set; }
        public int win { get; set; }
        public int loss { get; set; }
        public int draw { get; set; }
        public int points { get; set; }
        public object form { get; set; }
        public string link_url { get; set; }
        public int strength_overall_home { get; set; }
        public int strength_overall_away { get; set; }
        public int strength_attack_home { get; set; }
        public int strength_attack_away { get; set; }
        public int strength_defence_home { get; set; }
        public int strength_defence_away { get; set; }
        public int team_division { get; set; }
    }
}
