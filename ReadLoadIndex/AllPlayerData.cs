using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ReadLoadIndex
{
    public class AllPlayerData
    {
        public List<PlayerData> playerDatas = new List<PlayerData>();
        public Dictionary<int, string> allUnitDic = new Dictionary<int, string>();
        public List<(int, string)> dbList = new List<(int, string)>();
        public Dictionary<long, PlayerData> playerDic_diff;
    }
    public class PlayerData
    {
        public string name;
        public long view_id;
        public Dictionary<int, string[]> boxDic = new Dictionary<int, string[]>();
        public List<UnitDataS> unitList = new List<UnitDataS>();
    }
}
