using AutoRetainer.Internal;
using AutoRetainer.Modules.Voyage;
using AutoRetainerAPI.Configuration;
using Dalamud.Game.Text.SeStringHandling.Payloads;
using Dalamud.Utility;
using ECommons.ExcelServices;
using ECommons.ExcelServices.Sheets;
using ECommons.Interop;
using Lumina.Excel.Sheets;
using System.Collections.ObjectModel;
using System.Text.RegularExpressions;
using System.Windows.Forms.VisualStyles;
using GrandCompany = ECommons.ExcelServices.GrandCompany;

namespace AutoRetainer;

internal static class Lang
{
    internal const string CharPlant = "";
    internal const string CharLevel = "";
    internal const string CharItemLevel = "";
    internal const string CharDice = "";
    internal const string CharDeny = "";
    internal const string CharQuestion = "";
    internal const string CharLevelSync = "";
    internal const string CharP = "";
    internal const string StrDCV = "";

    internal const string IconRefresh = "\uf2f9";
    internal const string IconMultiMode = "\uf021";
    internal const string IconDuplicate = "\uf24d";
    internal const string IconGil = "\uf51e";
    internal const string IconPlanner = "\uf0ae";
    internal const string IconSettings = "\uf013";
    internal const string IconWarning = "\uf071";

    internal const string IconAnchor = "\uf13d";
    internal const string IconLevelup = "\ue098";
    internal const string IconResend = "\ue4bb";
    internal const string IconUnlock = "\uf13e";
    internal const string IconRepeat = "\uf363";
    internal const string IconPath = "\uf55b";
    internal const string IconFire = "\uf06d";

    internal static string Bool(bool? value) => value == null ? "未知" : value.Value ? "是" : "否";

    internal static string LogOutAndExitGame => Svc.Data.GetExcelSheet<Addon>().GetRow(116).Text.GetText(true).Cleanup();

    internal static readonly ReadOnlyDictionary<UnlockMode, string> UnlockModeNames = new(new Dictionary<UnlockMode, string>()
    {
        { UnlockMode.MultiSelect, "选择尽可能多的目的地" },
        { UnlockMode.SpamOne, "重复派遣单一目的地" },
        { UnlockMode.WhileLevelling, "升级时包含一个待解锁目的地" },
    });

    internal static readonly ReadOnlyDictionary<VesselBehavior, string> VesselBehaviorNames = new(new Dictionary<VesselBehavior, string>()
    {
        { VesselBehavior.Finalize, "仅结算奖励" },
        { VesselBehavior.Redeploy, "再次派遣" },
        { VesselBehavior.LevelUp, "最佳经验航线" },
        { VesselBehavior.Unlock, "解锁目的地" },
        { VesselBehavior.Use_plan, "使用目的地方案" },
    });

    internal static readonly ReadOnlyDictionary<PlanCompleteBehavior, string> PlanCompleteBehaviorNames = new(new Dictionary<PlanCompleteBehavior, string>()
    {
        { PlanCompleteBehavior.Restart_plan, "重新开始方案" },
        { PlanCompleteBehavior.Assign_Quick_Venture, "执行自由探索委托" },
        { PlanCompleteBehavior.Do_nothing, "不执行操作" },
        { PlanCompleteBehavior.Repeat_last_venture, "重复上次探险" },
    });

    internal static readonly ReadOnlyDictionary<UnavailableVentureDisplay, string> UnavailableVentureDisplayNames = new(new Dictionary<UnavailableVentureDisplay, string>()
    {
        { UnavailableVentureDisplay.Hide, "隐藏" },
        { UnavailableVentureDisplay.Display, "显示" },
        { UnavailableVentureDisplay.Allow_selection, "允许选择" },
    });

    internal static readonly ReadOnlyDictionary<MultiModeType, string> MultiModeTypeNames = new(new Dictionary<MultiModeType, string>()
    {
        { MultiModeType.Retainers, "仅雇员" },
        { MultiModeType.Submersibles, "仅远航探索" },
        { MultiModeType.Everything, "全部" },
    });

    internal static readonly ReadOnlyDictionary<GCDeliveryType, string> GCDeliveryTypeNames = new(new Dictionary<GCDeliveryType, string>()
    {
        { GCDeliveryType.Disabled, "禁用" },
        { GCDeliveryType.Hide_Armoury_Chest_Items, "隐去兵装库中的物品" },
        { GCDeliveryType.Hide_Gear_Set_Items, "隐去套装中的物品" },
        { GCDeliveryType.Show_All_Items, "显示所有物品" },
    });

    internal static readonly ReadOnlyDictionary<CutsceneSkipMode, string> CutsceneSkipModeNames = new(new Dictionary<CutsceneSkipMode, string>()
    {
        { CutsceneSkipMode.Never, "从不" },
        { CutsceneSkipMode.When_Multi_Mode_is_on, "仅多角色模式启用时" },
        { CutsceneSkipMode.Always, "始终" },
    });

    internal static readonly ReadOnlyDictionary<OpenBellBehavior, string> OpenBellBehaviorNames = new(new Dictionary<OpenBellBehavior, string>()
    {
        { OpenBellBehavior.Do_nothing, "不执行操作" },
        { OpenBellBehavior.Enable_AutoRetainer, "启用 AutoRetainer" },
        { OpenBellBehavior.Disable_AutoRetainer, "禁用 AutoRetainer" },
        { OpenBellBehavior.Pause_AutoRetainer, "暂停 AutoRetainer" },
    });

    internal static readonly ReadOnlyDictionary<TaskCompletedBehavior, string> TaskCompletedBehaviorNames = new(new Dictionary<TaskCompletedBehavior, string>()
    {
        { TaskCompletedBehavior.Close_retainer_list_and_disable_plugin, "关闭雇员列表并禁用插件" },
        { TaskCompletedBehavior.Close_retainer_list_and_keep_plugin_enabled, "关闭雇员列表并保持插件启用" },
        { TaskCompletedBehavior.Stay_in_retainer_list_and_disable_plugin, "停留在雇员列表并禁用插件" },
        { TaskCompletedBehavior.Stay_in_retainer_list_and_keep_plugin_enabled, "停留在雇员列表并保持插件启用" },
    });

    internal static readonly ReadOnlyDictionary<PluginEnableReason, string> PluginEnableReasonNames = new(new Dictionary<PluginEnableReason, string>()
    {
        { PluginEnableReason.Access, "访问传唤铃" },
        { PluginEnableReason.Manual, "手动" },
        { PluginEnableReason.Auto, "自动" },
        { PluginEnableReason.MultiMode, "多角色模式" },
        { PluginEnableReason.Artisan, "Artisan 联动" },
    });

    internal static readonly ReadOnlyDictionary<RetainersVisualOrder, string> RetainersVisualOrderNames = new(new Dictionary<RetainersVisualOrder, string>()
    {
        { RetainersVisualOrder.Ventures, "探险币" },
        { RetainersVisualOrder.Inventory_Slots, "背包空位" },
        { RetainersVisualOrder.Region_JP, "日本大区" },
        { RetainersVisualOrder.Region_NA, "北美大区" },
        { RetainersVisualOrder.Region_EU, "欧洲大区" },
        { RetainersVisualOrder.Region_OC, "大洋洲大区" },
        { RetainersVisualOrder.World, "服务器" },
        { RetainersVisualOrder.DataCenter, "数据中心" },
        { RetainersVisualOrder.Name, "名称" },
    });

    internal static readonly ReadOnlyDictionary<DeployablesVisualOrder, string> DeployablesVisualOrderNames = new(new Dictionary<DeployablesVisualOrder, string>()
    {
        { DeployablesVisualOrder.Ceruleum, "桶装青磷水" },
        { DeployablesVisualOrder.Repair_Kits, "魔导机械修理材料" },
        { DeployablesVisualOrder.Inventory_Slots, "背包空位" },
        { DeployablesVisualOrder.Region_JP, "日本大区" },
        { DeployablesVisualOrder.Region_NA, "北美大区" },
        { DeployablesVisualOrder.Region_EU, "欧洲大区" },
        { DeployablesVisualOrder.Region_OC, "大洋洲大区" },
        { DeployablesVisualOrder.World, "服务器" },
        { DeployablesVisualOrder.DataCenter, "数据中心" },
        { DeployablesVisualOrder.Name, "名称" },
    });

    internal static readonly ReadOnlyDictionary<ExcelWorldHelper.Region, string> RegionNames = new(new Dictionary<ExcelWorldHelper.Region, string>()
    {
        { ExcelWorldHelper.Region.JP, "日本大区" },
        { ExcelWorldHelper.Region.NA, "北美大区" },
        { ExcelWorldHelper.Region.EU, "欧洲大区" },
        { ExcelWorldHelper.Region.OC, "大洋洲大区" },
        { (ExcelWorldHelper.Region)5, "中国大区" },
        { (ExcelWorldHelper.Region)6, "韩国大区" },
        { (ExcelWorldHelper.Region)7, "北美云大区" },
    });

    internal static string RegionName(ExcelWorldHelper.Region region)
        => RegionNames.TryGetValue(region, out var name) ? name : $"未知大区（{(int)region}）";

    internal static readonly ReadOnlyDictionary<GCExchangeCategoryTab, string> GCExchangeCategoryNames = new(new Dictionary<GCExchangeCategoryTab, string>()
    {
        { GCExchangeCategoryTab.Weapons, "武具" },
        { GCExchangeCategoryTab.Armor, "防具" },
        { GCExchangeCategoryTab.Materiel, "军用品" },
        { GCExchangeCategoryTab.Materials, "展示品" },
    });

    internal static readonly ReadOnlyDictionary<ItemRarity, string> ItemRarityNames = new(new Dictionary<ItemRarity, string>()
    {
        { ItemRarity.White, "白色" },
        { ItemRarity.Green, "绿色" },
        { ItemRarity.Blue, "蓝色" },
        { ItemRarity.Purple, "紫色" },
        { ItemRarity.Pink, "粉色" },
    });

    internal static readonly ReadOnlyDictionary<GrandCompany, string> GrandCompanyNames = new(new Dictionary<GrandCompany, string>()
    {
        { GrandCompany.Unemployed, "未加入" },
        { GrandCompany.Maelstrom, "黑涡团" },
        { GrandCompany.TwinAdder, "双蛇党" },
        { GrandCompany.ImmortalFlames, "恒辉队" },
    });

    internal static readonly ReadOnlyDictionary<VoyageType, string> VoyageTypeNames = new(new Dictionary<VoyageType, string>()
    {
        { VoyageType.Airship, "飞空艇" },
        { VoyageType.Submersible, "潜水艇" },
    });

    internal static readonly ReadOnlyDictionary<PanelType, string> PanelTypeNames = new(new Dictionary<PanelType, string>()
    {
        { PanelType.None, "无" },
        { PanelType.Unknown, "未知" },
        { PanelType.TypeSelector, "类型选择" },
        { PanelType.Airship, "飞空艇" },
        { PanelType.Submersible, "潜水艇" },
    });

    private static string GetKeyName(LimitedKeys key)
    {
        var name = key.ToString();
        if(name.Length == 1 || (name.StartsWith('F') && int.TryParse(name.AsSpan(1), out _))) return name;
        if(name.StartsWith("Digit_")) return name[6..];
        if(name.StartsWith("NumPad")) return $"小键盘 {name[6..]}";
        return key switch
        {
            LimitedKeys.None => "未设置",
            LimitedKeys.LeftMouseButton => "鼠标左键",
            LimitedKeys.RightMouseButton => "鼠标右键",
            LimitedKeys.MiddleMouseButton => "鼠标中键",
            LimitedKeys.XMouseButton1 => "鼠标侧键 1",
            LimitedKeys.XMouseButton2 => "鼠标侧键 2",
            LimitedKeys.Back => "退格键",
            LimitedKeys.Tab => "Tab 键",
            LimitedKeys.Clear => "清除键",
            LimitedKeys.Enter => "回车键",
            LimitedKeys.Pause => "暂停键",
            LimitedKeys.CapsLock => "大写锁定键",
            LimitedKeys.Escape => "Esc 键",
            LimitedKeys.Space => "空格键",
            LimitedKeys.PageUp => "上一页键",
            LimitedKeys.PageDown => "下一页键",
            LimitedKeys.End => "End 键",
            LimitedKeys.Home => "Home 键",
            LimitedKeys.Left => "左方向键",
            LimitedKeys.Up => "上方向键",
            LimitedKeys.Right => "右方向键",
            LimitedKeys.Down => "下方向键",
            LimitedKeys.PrintScreen => "截图键",
            LimitedKeys.Insert => "插入键",
            LimitedKeys.Delete => "删除键",
            LimitedKeys.LeftShiftKey => "左 Shift 键",
            LimitedKeys.RightShiftKey => "右 Shift 键",
            LimitedKeys.LeftControlKey => "左 Ctrl 键",
            LimitedKeys.RightControlKey => "右 Ctrl 键",
            LimitedKeys.LeftAltKey => "左 Alt 键",
            LimitedKeys.RightAltKey => "右 Alt 键",
            _ => $"按键 {(int)key}",
        };
    }

    internal static readonly ReadOnlyDictionary<LimitedKeys, string> LimitedKeyNames = new(Enum.GetValues<LimitedKeys>().Distinct().ToDictionary(x => x, GetKeyName));

    private static readonly Dictionary<string, string> SubmarineClassNames = new()
    {
        { "Shark", "鲨鱼级" },
        { "Unkiu", "甲鲎级" },
        { "Whale", "须鲸级" },
        { "Coelacanth", "腔棘鱼级" },
        { "Syldra", "希尔德拉级" },
        { "ModShark", "鲨鱼改级" },
        { "ModUnkiu", "甲鲎改级" },
        { "ModWhale", "须鲸改级" },
        { "ModCoelacanth", "腔棘鱼改级" },
        { "ModSyldra", "希尔德拉改级" },
    };

    private static ReadOnlyDictionary<T, string> GetSubmarinePartNames<T>() where T : struct, Enum
        => new(Enum.GetValues<T>().ToDictionary(x => x, x => SubmarineClassNames[x.ToString()]));

    internal static readonly ReadOnlyDictionary<Hull, string> HullNames = GetSubmarinePartNames<Hull>();
    internal static readonly ReadOnlyDictionary<Stern, string> SternNames = GetSubmarinePartNames<Stern>();
    internal static readonly ReadOnlyDictionary<Bow, string> BowNames = GetSubmarinePartNames<Bow>();
    internal static readonly ReadOnlyDictionary<Bridge, string> BridgeNames = GetSubmarinePartNames<Bridge>();

    internal static readonly (string Normal, string GameFont) Digits = ("0123456789", "");

    //The summoning bell's own menu, in whatever language the client runs.
    private const string RetainerBellDialogue = "custom/000/CmnDefRetainerCall_00010";

    private static string BellText(uint row) => Svc.Data.GetExcelSheet<QuestDialogueText>(name: RetainerBellDialogue)?.GetRowOrDefault(row)?.Value.GetText() ?? "";

    //Everything a log message says before its first macro, which is fixed text in any language.
    internal static string LogMessageOpening(uint row)
    {
        var text = Svc.Data.GetExcelSheet<LogMessage>()?.GetRowOrDefault(row)?.Text.ToDalamudString();
        if(text == null) return "";
        var opening = "";
        foreach(var payload in text.Payloads)
        {
            if(payload is not TextPayload literal) break;
            opening += literal.Text;
        }
        return opening.Trim();
    }

    private static string AddonText(uint row) => Svc.Data.GetExcelSheet<Addon>()?.GetRowOrDefault(row)?.Text.GetText() ?? "";

    private static string EObjText(uint row) => Svc.Data.GetExcelSheet<EObjName>()?.GetRowOrDefault(row)?.Singular.GetText() ?? "";

    //Up to the first macro, as a menu entry may end with an amount.
    private static string DialogueText(string sheet, uint row) => Svc.Data.GetExcelSheet<QuestDialogueText>(name: sheet)?.GetRowOrDefault(row)?.Value.GetText(true) ?? "";

    //A text the client does not have would match every entry, so it is left out.
    private static string[] GameText(params string[] text) => [.. text.Where(x => x.Length > 0)];

    private const string AirshipDialogue = "custom/002/CmnDefCompanyCommanderBoard_00258";
    private const string SubmarineDialogue = "custom/004/HouFixCompanySubmarine_00447";

    private static readonly Dictionary<(string Sheet, uint Row), string[]> FixedParts = [];

    //What a string of the game says around its macros (item names, amounts), without whitespace.
    private static string[] GetFixedParts(string sheet, uint row)
    {
        if(FixedParts.TryGetValue((sheet, row), out var parts)) return parts;
        var text = sheet switch
        {
            nameof(Addon) => Svc.Data.GetExcelSheet<Addon>()?.GetRowOrDefault(row)?.Text,
            nameof(LogMessage) => Svc.Data.GetExcelSheet<LogMessage>()?.GetRowOrDefault(row)?.Text,
            nameof(Warp) => Svc.Data.GetExcelSheet<Warp>()?.GetRowOrDefault(row)?.Question,
            _ => Svc.Data.GetExcelSheet<QuestDialogueText>(name: sheet)?.GetRowOrDefault(row)?.Value,
        };
        parts = text == null ? [] : [.. text.Value.ToDalamudString().Payloads.OfType<TextPayload>().Select(x => x.Text.Cleanup()).Where(x => x.Length > 0)];
        FixedParts[(sheet, row)] = parts;
        return parts;
    }

    //Whether the text on screen is one of these rows with the macros filled in: every fixed part of the row, in order.
    private static bool IsGameText(string text, string sheet, params uint[] rows)
    {
        text = text.Cleanup();
        foreach(var row in rows)
        {
            var parts = GetFixedParts(sheet, row);
            if(parts.Length == 0) continue;
            var at = 0;
            foreach(var part in parts)
            {
                at = text.IndexOf(part, at, StringComparison.OrdinalIgnoreCase);
                if(at < 0) break;
                at += part.Length;
            }
            if(at >= 0) return true;
        }
        return false;
    }

    internal static bool IsAddon(string text, params uint[] rows) => IsGameText(text, nameof(Addon), rows);
    internal static bool IsLogMessage(string text, params uint[] rows) => IsGameText(text, nameof(LogMessage), rows);
    internal static bool IsPanelAirship(string text) => IsGameText(text, AirshipDialogue, 1)
        || text.ContainsAny(StringComparison.OrdinalIgnoreCase, "請選擇飛空艇。", "请选择飞空艇。");
    internal static bool IsPanelSubmersible(string text) => IsGameText(text, SubmarineDialogue, 2)
        || text.ContainsAny(StringComparison.OrdinalIgnoreCase, "請選擇潛水艇。", "请选择潜水艇。");
    internal static bool IsHouseEntrance(string text) => IsGameText(text, nameof(Warp), 131148);

    private static string[] FuelNames => field ??= GetFuelNames();

    private static string[] GetFuelNames()
    {
        var item = Svc.Data.GetExcelSheet<Item>()?.GetRowOrDefault(10155);
        return GameText(item?.Singular.GetText() ?? "", item?.Plural.GetText() ?? "", item?.Name.GetText() ?? "");
    }

    internal static bool FuelPurchaseConfirm(string text) => IsAddon(text, 3473)
        && FuelNames.Any(x => text.Contains(x, StringComparison.OrdinalIgnoreCase));

    //196	TASK_CATEGORY_TREASURE	Field Exploration.
    //198	TASK_CATEGORY_MINER_2	Highland Exploration.
    //200	TASK_CATEGORY_BOTANIST_2	Woodland Exploration.
    //202	TASK_CATEGORY_FISHER_2	Waterside Exploration.
    internal static string[] FieldExplorationNames => field ??=
    [
        BellText(196),
        BellText(198),
        BellText(200),
        BellText(202),
        //zh-TW: row ids unverified on that client, so it stays a literal
        "平地探索委託（需要2枚探險幣）",
        "山岳探索委託（需要2枚探險幣）",
        "森林探索委託（需要2枚探險幣）",
        "水岸探索委託（需要2枚探險幣）",
    ];

    //195	TASK_CATEGORY_NORMAL	Hunting.
    //197	TASK_CATEGORY_MINER_1	Mining.
    //199	TASK_CATEGORY_BOTANIST_1	Botany.
    //201	TASK_CATEGORY_FISHER_1	Fishing.
    internal static string[] HuntingVentureNames => field ??=
    [
        BellText(195),
        BellText(197),
        BellText(199),
        BellText(201),
        //zh-TW: row ids unverified on that client, so it stays a literal
        "狩獵籌集委託（需要1枚探險幣）",
        "採礦籌集委託（需要1枚探險幣）",
        "採伐籌集委託（需要1枚探險幣）",
        "捕魚籌集委託（需要1枚探險幣）",
    ];

    //402	TASK_CATEGORY_FORTUNE	Quick Exploration.
    internal static string[] QuickExploration => field ??= [BellText(402), /*zh-TW: row id unverified, stays a literal*/ "自由尋寶委託（需要2枚探險幣）"];

    internal static string[] Entrance => field ??= GameText(EObjText(2002737), EObjText(2002739), "進入房屋", "进入房屋");

    internal static string ApartmentEntrance => Svc.Data.GetExcelSheet<EObjName>().GetRow(2007402).Singular.ToString();

    internal static bool ConfirmHouseEntrance(string text) => IsHouseEntrance(text)
        || text.ContainsAny(StringComparison.OrdinalIgnoreCase, "要進入這間房屋嗎？", "要进入这间房屋吗？");

    //194	ASK_CATEGORY	Select a category.
    internal static string[] RetainerAskCategoryText => field ??= [BellText(194), /*zh-TW: row id unverified, stays a literal*/ "請選擇要委託的探險"];

    internal static string[] BellName => [Svc.Data.GetExcelSheet<EObjName>().GetRow(2000401).Singular.GetText(), "リテイナーベル"];

    //0	TEXT_HOUFIXMANSIONENTRANCE_00359_HOUSINGAREA_MENU_ENTER_MYROOM	Go to your apartment
    //0	TEXT_HOUFIXMANSIONENTRANCE_00359_HOUSINGAREA_MENU_ENTER_MYROOM	自分の部屋に移動する
    //0	TEXT_HOUFIXMANSIONENTRANCE_00359_HOUSINGAREA_MENU_ENTER_MYROOM	Die eigene Wohnung betreten
    //0	TEXT_HOUFIXMANSIONENTRANCE_00359_HOUSINGAREA_MENU_ENTER_MYROOM	Aller dans votre appartement

    internal static string[] GoToYourApartment => field ??= GameText(
        DialogueText("custom/003/HouFixMansionEntrance_00359", 0),
        DialogueText("custom/003/HouFixMansionEntrance_00359", 4), "移動到自己的房間", "移动到自己的房间");

    internal static string[] SkipCutsceneStr => field ??= GameText(AddonText(281), AddonText(5882), "要跳過這段過場動畫嗎？", "要跳过这段过场动画吗？");
    //11	TEXT_CMNDEFHOUSINGPERSONALROOMENTRANCE_00178_GOTO_WORKSHOP	Move to the company workshop
    //11	TEXT_CMNDEFHOUSINGPERSONALROOMENTRANCE_00178_GOTO_WORKSHOP	地下工房に移動する
    //11	TEXT_CMNDEFHOUSINGPERSONALROOMENTRANCE_00178_GOTO_WORKSHOP	Die Ge<SoftHyphen/>sell<SoftHyphen/>schaftswerkstätte betreten
    //11	TEXT_CMNDEFHOUSINGPERSONALROOMENTRANCE_00178_GOTO_WORKSHOP	Aller dans l'atelier de compagnie
    internal static string[] EnterWorkshop => field ??= GameText(DialogueText("custom/001/CmnDefHousingPersonalRoomEntrance_00178", 11), "移動到部隊工房", "移动到部队工房");

    internal static string[] AirshipManagement => field ??= GameText(
        Svc.Data.GetExcelSheet<CustomTalk>()?.GetRowOrDefault(721154)?.MainOption.GetText() ?? "", "管理飛空艇", "管理飞空艇");
    internal static string[] SubmarineManagement => field ??= GameText(
        Svc.Data.GetExcelSheet<CustomTalk>()?.GetRowOrDefault(721343)?.MainOption.GetText() ?? "", "管理潛水艇", "管理潜水艇");
    internal static string[] CancelVoyage => field ??= GameText(AddonText(2), "取消");
    internal static string[] NothingVoyage => field ??= GameText(DialogueText(SubmarineDialogue, 6), DialogueText(AirshipDialogue, 5), "取消");
    internal static string[] DeployOnSubaquaticVoyage => field ??= GameText(DialogueText(SubmarineDialogue, 8), "出發", "出发");
    internal static string[] ViewPrevVoyageLog => field ??= GameText(DialogueText(SubmarineDialogue, 9), DialogueText(AirshipDialogue, 8), "上次的遠航報告", "上次的远航报告");
    internal static string[] VoyageQuitEntry => field ??= GameText(AddonText(5), "取消");
    internal static string[] ChangeSubmersibleComponents => field ??= GameText(DialogueText(SubmarineDialogue, 10), "更換配件", "更换配件");
    internal static bool RegisterSub(string text) => IsGameText(text, SubmarineDialogue, 3)
        || text.StartsWith("登記新的潛水艇", StringComparison.Ordinal)
        || text.StartsWith("登记新的潜水艇", StringComparison.Ordinal);
    //2004353	entrance to additional chambers	0	entrances to additional chambers	0	1	1	0	0
    internal static string[] AdditionalChambersEntrance =>
    [
        Svc.Data.GetExcelSheet<EObjName>().GetRow(2004353).Singular.GetText(),
        Regex.Replace(Svc.Data.GetExcelSheet<EObjName>().GetRow(2004353).Singular.GetText(), @"\[.*?\]", "")
    ];

    //2005274	voyage control panel	0	voyage control panels	0	0	1	0	0
    internal static string PanelName => Svc.Data.GetExcelSheet<EObjName>().GetRow(2005274).Singular.GetText();

    //4160	60	9	0	False	Unable to retrieve extracted items. Insufficient inventory/crystal inventory space.
    internal static string VoyageInventoryError => Svc.Data.GetExcelSheet<LogMessage>().GetRow(4160).Text.ToDalamudString().GetText();

    internal static bool UnableToVisitWorld(string text) => IsLogMessage(text, 5800)
        || text.StartsWith("无法进行该操作，角色正在访问其他数据中心", StringComparison.Ordinal)
        || text.StartsWith("無法進行該操作，角色正在訪問其他資料中心", StringComparison.Ordinal)
        || (text.StartsWith("角色正在", StringComparison.Ordinal) && text.EndsWith("大区游玩，无法操作。", StringComparison.Ordinal))
        || (text.StartsWith("由於正前往", StringComparison.Ordinal) && text.EndsWith("DC遊玩，無法操作。", StringComparison.Ordinal));

    //4169	60	9	0	False	Unable to repair vessel component without the required <SheetEn(Item,3,IntegerParameter(1),1,1)/>.
    //4272	60	9	0	False Unable to repair vessel.Insufficient<SheetEn(Item,3,IntegerParameter(1),3,1)/>.
    //4169	60	9	0	False	修理に必要な<Sheet(Item,IntegerParameter(1),0)/>を持っていません。
    //4272	60	9	0	False	修理に必要な<Sheet(Item,IntegerParameter(1),0)/>が足りません。
    //4169	60	9	0	False	未持有修理所必需的<Sheet(Item,IntegerParameter(1),0)/>。
    //4272	60	9	0	False	沒有修理所必需的<Sheet(Item,IntegerParameter(1),0)/>。
    //4272	60	9	0	False	Du hast nicht genug <SheetDe(Item,5,IntegerParameter(1),2,4,1)/> für die Reparatur.
    //4169	60	9	0	False	Für die Reparatur ist <SheetDe(Item,1,IntegerParameter(1),1,1,1)/> erforderlich.
    //4169	60	9	0	False	Réparation impossible. Vous n'avez pas <SheetFr(Item,2,IntegerParameter(1),1,1)/> nécessaire.
    //4272	60	9	0	False	Vous n'avez pas <SheetFr(Item,2,IntegerParameter(1),1,1)/> nécessaire à la réparation.

    internal static bool UnableToRepairVessel(string text) => IsLogMessage(text, 4169, 4272)
        || text.Contains("修理所必需的", StringComparison.OrdinalIgnoreCase);

    //11	TEXT_HOUFIXCOMPANYSUBMARINE_00447_SUBMARINE_CMD_REPAIR_PARTS	パーツの修理
    //11	TEXT_HOUFIXCOMPANYSUBMARINE_00447_SUBMARINE_CMD_REPAIR_PARTS	Bauteile reparieren
    //11	TEXT_HOUFIXCOMPANYSUBMARINE_00447_SUBMARINE_CMD_REPAIR_PARTS	Réparer des éléments
    //11	TEXT_HOUFIXCOMPANYSUBMARINE_00447_SUBMARINE_CMD_REPAIR_PARTS	修理配件
    //10	TEXT_CMNDEFCOMPANYCOMMANDERBOARD_00258_AIRSHIP_CMD_REPAIR_PARTS	パーツの修理
    //10	TEXT_CMNDEFCOMPANYCOMMANDERBOARD_00258_AIRSHIP_CMD_REPAIR_PARTS	Bauteile reparieren
    //10	TEXT_CMNDEFCOMPANYCOMMANDERBOARD_00258_AIRSHIP_CMD_REPAIR_PARTS	Réparer des éléments
    //10	TEXT_CMNDEFCOMPANYCOMMANDERBOARD_00258_AIRSHIP_CMD_REPAIR_PARTS	修理配件

    internal static string[] WorkshopRepair => field ??= GameText(DialogueText(SubmarineDialogue, 11), DialogueText(AirshipDialogue, 10), "修理配件");

    //Use <If(Equal(IntegerParameter(4),1))>your last <SheetEn(Item,3,IntegerParameter(2),1,1)/><Else/><Value>IntegerParameter(3)</Value> of your <Value>IntegerParameter(4)</Value> <SheetEn(Item,3,IntegerParameter(2),2,1)/></If> to repair your vessel's <SheetEn(Item,3,IntegerParameter(1),1,1)/>?
    //6587	<If(Equal(IntegerParameter(3),1))><Clickable(<SheetDe(Item,2,IntegerParameter(2),1,4,1)/>)/><Else/><Value>IntegerParameter(3)</Value> <SheetDe(Item,5,IntegerParameter(2),2,4,1)/></If> (Besitz: <Value>IntegerParameter(4)</Value>) benutzen, um <SheetDe(Item,2,IntegerParameter(1),1,4,1)/> zu reparieren?
    //6587	Utiliser <If(Equal(IntegerParameter(3),1))><SheetFr(Item,1,IntegerParameter(2),1,1)/><Else/><Value>IntegerParameter(3)</Value> <SheetFr(Item,12,IntegerParameter(2),2,1)/></If> pour réparer <SheetFr(Item,2,IntegerParameter(1),1,1)/> de votre appareil<Indent/>? (<Value>IntegerParameter(4)</Value> possédé<If(LessThanOrEqualTo(IntegerParameter(4),1))><Else/>s</If>)
    /*6587	下記のアイテムを修理しますか？
    <Sheet(Item,IntegerParameter(1),0)/>
    消費:<Sheet(Item,IntegerParameter(2),0)/>×<Value>IntegerParameter(3)</Value>(所持数 <Value>IntegerParameter(4)</Value>)
    */

    internal static bool WorkshopRepairConfirm(string text) => IsAddon(text, 6587)
        || text.ContainsAny(StringComparison.OrdinalIgnoreCase, "要修理下列部件嗎", "要修理下列元件嗎", "要修理下列組件嗎", "要修理下列部件吗");

    // Use the components selected and <If(Equal(IntegerParameter(1),1))>the following item<Else/><Value>IntegerParameter(1)</Value> of the following items</If> to outfit and register your submersible?
    /* 6886 Das Tauchboot mit den gewählten Bauteilen registrieren?
     Verbraucht <Value>IntegerParameter(1)</Value> <If(Equal(IntegerParameter(1),1))>Exemplar<Else/>Exemplare</If> des folgenden Gegenstands:
    */
    // 6886 Utiliser les éléments choisis et <If(Equal(IntegerParameter(1),1))>l'objet suivant<Else/><Value>IntegerParameter(1)</Value> des objets suivants</If> pour équiper et enregistrer le sous-marin<Indent/>?
    /*選択したパーツアイテムと以下のアイテムを
       <Value>IntegerParameter(1)</Value>枚消費して潜水艦を登録します。
       よろしいですか？
    */

    internal static bool WorkshopRegisterConfirm(string text) => IsAddon(text, 6886)
        || text.ContainsAny(StringComparison.OrdinalIgnoreCase, "張下列道具登記新的潛水艇嗎", "张下列道具登记新的潜水艇吗",
            "消耗下列道具登记新的潜水艇吗", "消耗下列道具登記新的潛水艇嗎");

    //Your retainer will be unable to process item buyback requests once recalled. Are you sure you wish to proceed?
    //215	TEXT_CMNDEFRETAINERCALL_00010_ASK_RETURN_WITH_BUYBACK	Wenn du deinen Gehilfen wegschickst, kannst du die von ihm verkauften Gegenstände nicht mehr zurückkaufen. Möchtest du trotzdem fortfahren?
    //215	TEXT_CMNDEFRETAINERCALL_00010_ASK_RETURN_WITH_BUYBACK	Renvoyer le servant effacera la liste de rachat. Confirmer<Indent/>?

    //215	ASK_RETURN_WITH_BUYBACK	Your retainer will be unable to process item buyback requests once recalled.
    internal static string[] WillBeUnableToProcessBuyback => field ??= [BellText(215)];

    //3290	<Sheet(Item,IntegerParameter(1),0)/>×<Value>IntegerParameter(2)</Value>を、<Format(IntegerParameter(3),FF022C)/>枚の軍票と交換します。
    //よろしいですか？
    //3290	<Format(IntegerParameter(3),FF022E)/> Staatstaler gegen <If(Equal(IntegerParameter(2),1))><SheetDe(Item,1,IntegerParameter(1),1,4,1)/><Else/><Format(IntegerParameter(2),FF022E)/> <SheetDe(Item,5,IntegerParameter(1),2,4,1)/></If> eintauschen?
    //3290	Acheter <Value>IntegerParameter(2)</Value> <SheetFr(Item,12,IntegerParameter(1),IntegerParameter(2),1)/> pour <Format(IntegerParameter(3),FF05021D0103)/> sceau<If(LessThanOrEqualTo(IntegerParameter(3),1))><Else/>x</If><Indent/>?

    internal static bool GCSealExchangeConfirm(string text) => IsAddon(text, 3290)
        || text.ContainsAny(StringComparison.OrdinalIgnoreCase, "要交換嗎", "要交换吗");

    internal static bool DiscardItem(string text) => IsAddon(text, 110, 153)
        || (text.StartsWith("確定要捨棄", StringComparison.Ordinal) && text.EndsWith("嗎？", StringComparison.Ordinal))
        || (text.StartsWith("确定要舍弃", StringComparison.Ordinal) && text.EndsWith("吗？", StringComparison.Ordinal));
}
