using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ReadLoadIndex
{
    public class LoadData
    {
        public DataHead data_headers;
        public LoadDataBody data;
    }
    public class DataHead
    {
        public string request_id = "1234564894";
        public int short_udid = 0;
        public long viewer_id = 0;
        public string sid = "000";
        public long servertime = 1615604550;
        public int result_code;
        public DataHead() { }
        public DataHead(int resultcode, long time = 0)
        {
            servertime = time;
            result_code = resultcode;
        }
    }
    public class LoadDataBody
    {
        public UserInfo user_info;
        public object user_jewel;
        public object user_gold;
        public List<UnitDataS> unit_list;
        public List<UserChara> user_chara_info;
        //public List<Deck> deck_list;
        public object item_list;
        public object user_equip;
        public object shop;
        public object ini_setting;
        public int max_storage_num;
        public List<int> campaign_list;
        public int can_free_gacha;
        public int can_campaign_gacha;
        public object gacha_point_info;
        public List<int> read_story_ids;
        public List<string> unlock_story_ids;
        public object event_statuses;
        public object drj;
        public List<UserMyParty> user_my_party;
        public List<UserMyPartyTab> user_my_party_tab;
        public long daily_reset_time;
        public object login_bonus_list;
        public int present_count;
        public int clan_like_count;
        public object dispatch_units;
        public object clan_battle;
        public object voice;
        public long qr;
        public int ddn;
        public long pa;
        //public long as;
        public long gs;
        public long tq;
        public long sv;
        public long jb;
        public long tf;
        public long ue;
        public long em;
        public long bv;
        public long gl;
        public long ccb;
        public long pfm;
        public long tt;
        //public List<MyPage> my_page;
    }
    public class UserInfo
    {
        public long viewer_id;
        public string user_name;
        public string user_comment;
        public int team_level;
        public int user_stamina;
        public int max_stamina;
        public int team_exp;
        public int favorite_unit_id;
        public int tutorial_flag;
        public int invite_accept_flag;
        public int user_birth;
        public int platform_id;
        public int channel_id;
        public string last_ac;
        public int last_ac_time;
        public int server_id;
        public int reg_time;
        public int stamina_full_recovery_time;
        public object emblem;
    }
    public class UserGood
    {
        public int gold_id_pay = 0;
        public int gold_id_free = 9999999;
    }

    public class UserMyParty
    {
        public int tab_number;
        public int party_number;
        public int party_label_type;
        public string party_name;
        public int unit_id_1;
        public int unit_id_2;
        public int unit_id_3;
        public int unit_id_4;
        public int unit_id_5;
        public UserMyParty() { }
        public UserMyParty(int tab_number, int party_number, int party_label_type, string party_name, int unit_id_1, int unit_id_2, int unit_id_3, int unit_id_4, int unit_id_5)
        {
            this.tab_number = tab_number;
            this.party_number = party_number;
            this.party_label_type = party_label_type;
            this.party_name = party_name;
            this.unit_id_1 = unit_id_1;
            this.unit_id_2 = unit_id_2;
            this.unit_id_3 = unit_id_3;
            this.unit_id_4 = unit_id_4;
            this.unit_id_5 = unit_id_5;
        }
        public bool AllZero()
        {
            return unit_id_1 + unit_id_2 + unit_id_3 + unit_id_4 + unit_id_5 == 0;
        }
        public List<int> GetUnits()
        {
            List<int> result = new List<int>();
            if (unit_id_1 > 0)
                result.Add(unit_id_1);
            if (unit_id_2 > 0)
                result.Add(unit_id_2);
            if (unit_id_3 > 0)
                result.Add(unit_id_3);
            if (unit_id_4 > 0)
                result.Add(unit_id_4);
            if (unit_id_5 > 0)
                result.Add(unit_id_5);
            return result;
        }
    }
    public class UserMyPartyTab
    {
        public int tab_number;
        public string tab_name;
        public UserMyPartyTab() { }

        public UserMyPartyTab(int tab_number, string tab_name)
        {
            this.tab_number = tab_number;
            this.tab_name = tab_name;
        }
    }
    public class MyPage
    {
        public int type;
        public int id;
        public int music_id;
        public int still_skin_id;

        public MyPage() { }
        public MyPage(int type, int id, int music_id, int still_skin_id)
        {
            this.type = type;
            this.id = id;
            this.music_id = music_id;
            this.still_skin_id = still_skin_id;
        }
    }

    public class UnitDataS
    {
        public int id;

        public int get_time;

        public int unit_rarity;
        public int battle_rarity = 3;
        public int unit_level;

        public int unit_exp;
        public int promotion_level;

        public List<SkillLevelInfo> union_burst;

        public List<SkillLevelInfo> main_skill;

        public List<SkillLevelInfo> ex_skill;

        public List<SkillLevelInfo> free_skill;
        public List<EquipSlot> equip_slot;

        public List<EquipSlot> unique_equip_slot;
        public int power = -1;
        //public SkinData skin_data;
        public Unlock6Item unlock_rarity_6_item = new Unlock6Item(0);
        public UnitDataS()
        {

        }


        public string GetEquipList()
        {
            var Result = new string[6] { "-", "-", "-", "-", "-", "-" };
            for (int i = 0; i < 6; i++)
            {
                if (equip_slot.Count > i)
                {
                    int lv = equip_slot[i].GetLv();
                    Result[i] = lv >= 0 ? lv.ToString() : "-";
                }
            }
            return $"{Result[0]}{Result[1]}{Result[2]}{Result[3]}{Result[4]}{Result[5]}";
        }
        public int[] GetEquipListInt()
        {
            var Result = new int[6] { -1,-1,-1,-1,-1,-1 };
            for (int i = 0; i < 6; i++)
            {
                if (equip_slot.Count > i)
                {
                    int lv = equip_slot[i].GetLv();
                    Result[i] = lv;
                }
            }
            return Result;
        }
        public int[] GetSkillLevelInfo()
        {
            int[] result = new int[4] { 0, 0, 0, 0 };
            if (union_burst != null && union_burst.Count > 0)
                result[0] = (union_burst[0].skill_level);
            if (main_skill != null && main_skill.Count > 0)
                result[1] = (main_skill[0].skill_level);
            if (main_skill != null && main_skill.Count > 1)
                result[2] = (main_skill[1].skill_level);
            if (ex_skill != null && ex_skill.Count > 0)
                result[3] = (ex_skill[0].skill_level);
            return result;
        }
        public int GetUek()
        {
            if (unique_equip_slot != null && unique_equip_slot.Count > 0)
                return unique_equip_slot[0].GetLv();
            return 0;
        }
        /// <summary>
        /// 比较对象
        /// </summary>
        /// <param name="diff"></param>
        /// <returns>0-没改过，>0-绿色改动，<0-红色改动</returns>
        public int[] CopmairAll(UnitDataS diff)
        {
            int[] result = new int[7];
            result[0] = diff.unit_level - unit_level;
            result[1] = diff.unit_rarity == unit_rarity ? 0 : 1;
            result[2] = 0;
            result[3] = diff.promotion_level - promotion_level;

            result[4] = CopmairEquip(diff);
            result[5] = CompairSkill(diff);
            result[6] = diff.GetUek() - GetUek();
            return result;
        }
        private int CopmairEquip(UnitDataS diff)
        {
            if(diff.promotion_level > promotion_level)
            {
                return 1;
            }
            else if(diff.promotion_level == promotion_level)
            {
                int[] old1 = GetEquipListInt();
                int[] new1 = diff.GetEquipListInt();
                bool changed = false;
                bool isok = true;
                for(int i = 0; i < old1.Length; i++)
                {
                    if (new1[i] > old1[i])
                        changed = true;
                    if (new1[i] < old1[i])
                        isok = false;

                }
                return isok ? (changed ? 1 : 0) : -1;
            }
            return -1;
        }
        private int CompairSkill(UnitDataS diff)
        {
            int[] old1 = GetSkillLevelInfo();
            int[] new1 = diff.GetSkillLevelInfo();
            bool changed = false;
            bool isok = true;
            for (int i = 0; i < old1.Length; i++)
            {
                if (new1[i] > old1[i])
                    changed = true;
                if (new1[i] < old1[i])
                    isok = false;

            }
            return isok ? (changed ? 1 : 0) : -1;

        }
    }
    public class Unlock6Item
    {
        public int quest_clear = 0;
        public int slot_1 = 0;
        public int slot_2 = 0;
        public int slot_3 = 0;
        public Unlock6Item() { }
        public Unlock6Item(int a)
        {

        }
    }
    public class SkillLevelInfo
    {
        public int skill_level;
        public int skill_id;
        public SkillLevelInfo()
        {

        }
        public SkillLevelInfo(int skill_level, int skill_id)
        {
            this.skill_level = skill_level;
            this.skill_id = skill_id;
        }
    }
    public class EquipSlot
    {
        public int id;
        public int is_slot;
        public int enhancement_level;
        public int enhancement_pt;
        //public int rank;
        public EquipSlot()
        {

        }

        public int GetLv()
        {
            if (is_slot > 0)
                return enhancement_level;
            return -1;

        }

    }
    public class UserChara
    {
        public int chara_id;
        public int chara_love;//0-175(175)-245(420)-280(700)-700(1400)-700(2100)-700(2800)-1400(4200)
        public int love_level;
        public static int[] LOVE_VALUE = new int[20]
        {
            0, 0, 175, 420, 700, 1400, 2100, 2800, 4200,4255,
            4500,4550,4580,4750,4987,8521,9000,9002,9500,9999
        };
        public UserChara()
        {

        }
        public UserChara(int unitid, int loveLevel)
        {
            if (loveLevel <= 0)
                loveLevel = 1;
            chara_id = (int)Math.Round(unitid / 100.0f);
            //chara_id = unitid;
            chara_love = LOVE_VALUE[loveLevel];
            love_level = loveLevel;
        }
        public void Update(int loveLevel)
        {
            if (loveLevel <= 0)
                loveLevel = 1;
            chara_love = LOVE_VALUE[loveLevel];
            love_level = loveLevel;

        }
    }
    public class UnitStoryData
    {
        public int unitid;
        public List<int> stateStories = new List<int>();
        public int GetStoryLove(List<int> readStories)
        {
            switch (stateStories.Count)
            {
                case 7:
                case 11:
                    for(int i= stateStories.Count-1; i>=0; i--)
                    {
                        if (readStories.Contains(stateStories[i]))
                        {
                            return i + 2;
                        }
                    }
                    return 0;
                case 3:
                    if (readStories.Contains(stateStories[2]))
                    {
                        return 8;
                    }
                    if (readStories.Contains(stateStories[1]))
                    {
                        return 5;
                    }
                    if (readStories.Contains(stateStories[0]))
                    {
                        return 4;
                    }
                    return 0;
                default:
                    return -1;
            }
        }
    }
}
