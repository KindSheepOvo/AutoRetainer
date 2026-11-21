using ECommons.ExcelServices;
using FFXIVClientStructs.FFXIV.Client.Game;
using FFXIVClientStructs.FFXIV.Client.Game.UI;
using FFXIVClientStructs.FFXIV.Client.UI;
using FFXIVClientStructs.FFXIV.Client.UI.Agent;
using FFXIVClientStructs.FFXIV.Client.UI.Misc;
using Lumina.Excel.Sheets;
using System;
using System.Collections.Generic;
using System.Text;
using TerraFX.Interop.Windows;
using Cabinet = Lumina.Excel.Sheets.Cabinet;

namespace AutoRetainer.UI.NeoUI.AdvancedEntries.DebugSection;

public unsafe class DebugCabinet : DebugSectionBase
{
    public override void Draw()
    {
        ImGuiEx.TextCopy($"{((nint)(&UIState.Instance()->Cabinet) + 4):X}");
        ImGuiEx.Text($"""

            状态 1： {*(int*)((nint)(&UIState.Instance()->Cabinet))}
            状态 2： {*(int*)((nint)(&UIState.Instance()->Cabinet) + 4)}
            """);
        if(ImGuiEx.Button("执行命令 423", ImGuiEx.Ctrl))
        {
            GameMain.ExecuteCommand(423);
        }
        ref var onlyIn = ref Ref<bool>.Get("OnlyIn");
        if(ImGui.CollapsingHeader("收藏柜物品"))
        {
            ImGui.Checkbox("显示全部##1", ref onlyIn);
            var state = UIState.Instance()->Cabinet.State;
            if(state == FFXIVClientStructs.FFXIV.Client.Game.UI.Cabinet.CabinetState.Loaded)
            {
                if(ImGuiEx.BeginDefaultTable("CabItems", ["~物品", "收藏柜 ID", "已存入收藏柜"]))
                {
                    foreach(var x in Cabinet.Values)
                    {
                        var cab = UIState.Instance()->Cabinet.IsItemInCabinet(x.RowId);
                        if(!onlyIn && !cab) continue;
                        ImGui.TableNextRow();
                        ImGui.TableNextColumn();
                        ImGuiEx.Text($"{ExcelItemHelper.GetName(x.Item.RowId, true)}");
                        ImGui.TableNextColumn();
                        ImGuiEx.Text($"{x.RowId}");
                        ImGui.TableNextColumn();
                        ImGuiEx.Text(cab ? EColor.GreenBright : null, cab.ToString());
                    }
                    ImGui.EndTable();
                }
            }
            else
            {
                ImGuiEx.Text("未加载");
            }
        }
        if(ImGui.CollapsingHeader("收藏柜物品（含缓存）"))
        {
            ImGui.Checkbox("仅显示已存入物品##2", ref onlyIn);
            if(S.CabinetManager.TryGetStoredCabinetItems(out bool itemsCached, out var cabItems))
            {
                ImGuiEx.Text($"已缓存：{Lang.Bool(itemsCached)}");
                if(ImGuiEx.BeginDefaultTable("CabItems", ["~物品", "收藏柜 ID", "AR 判定"]))
                {
                    foreach(var x in Cabinet.Values)
                    {
                        var cab = cabItems.Contains(x.Item.RowId);
                        if(!onlyIn && !cab) continue;
                        ImGui.TableNextRow();
                        ImGui.TableNextColumn();
                        ImGuiEx.Text($"{ExcelItemHelper.GetName(x.Item.RowId, true)}");
                        ImGui.TableNextColumn();
                        ImGuiEx.Text($"{x.RowId}");
                        ImGui.TableNextColumn();
                        ImGuiEx.Text(cab ? EColor.GreenBright : null, cab.ToString());
                    }
                    ImGui.EndTable();
                }
            }
        }
        if(ImGui.CollapsingHeader("收藏柜物品（仅离线数据）"))
        {
            ImGui.Checkbox("仅显示已存入物品", ref onlyIn);
            if(ImGui.Button("写入离线数据")) OfflineDataManager.WriteOfflineData(true, true);
            if(ImGuiEx.BeginDefaultTable("CabItems", ["~物品", "收藏柜 ID", "AR 判定"]))
            {
                foreach(var x in Cabinet.Values)
                {
                    var cab = Data.StoredCabinetItems.Contains(x.Item.RowId);
                    if(!onlyIn && !cab) continue;
                    ImGui.TableNextRow();
                    ImGui.TableNextColumn();
                    ImGuiEx.Text($"{ExcelItemHelper.GetName(x.Item.RowId, true)}");
                    ImGui.TableNextColumn();
                    ImGuiEx.Text($"{x.RowId}");
                    ImGui.TableNextColumn();
                    ImGuiEx.Text(cab ? EColor.GreenBright : null, cab.ToString());
                }
                ImGui.EndTable();
            }
        }
        if(ImGui.CollapsingHeader("自动存入幻化台"))
        {
            if(ImGui.Button("前往旅馆并存入幻化台"))
            {
                S.MirageManager.EnqueueGoToInnAndDeliverEverything();
            }
            foreach(var inv in Utils.PlayerEntireInventory)
            {
                if(ImGui.CollapsingHeader($"{inv}"))
                {
                    ImGui.Indent();
                    var invCont = InventoryManager.Instance()->GetInventoryContainer(inv);
                    for(int i = 0; i < invCont->Size; i++)
                    {
                        var item = invCont->Items[i];
                        if(item.ItemId != 0 && Svc.Data.GetExcelSheet<Item>().TryGetRow(item.ItemId, out var data) && Utils.CanItemBePartOfMirageSet(item.ItemId))
                        {
                            ImGuiEx.Text(Utils.IsItemAlreadyMiraged(item.ItemId) ? EColor.GreenBright : null, $"{data.RowId} {data.Name}");
                        }
                    }
                    ImGui.Unindent();
                }
            }
        }
        ImGui.Separator();
        ImGuiEx.Text($"可存入收藏柜：{Lang.Bool(S.CabinetManager.CanDeliverCabinet())}");
        if(ImGui.Button("存入物品")) S.CabinetManager.EnqueueAllDeliverableItems();
        if(ImGui.Button("前往旅馆并存入全部物品")) S.CabinetManager.EnqueueGoToInnAndDeliverEverything();
        if(S.CabinetManager.TryGetStoredCabinetItems(out var cached, out var items))
        {
            ImGuiEx.Text($"已缓存：{Lang.Bool(cached)}");
            if(ImGui.CollapsingHeader("物品"))
            {
                foreach(var x in items)
                {
                    ImGuiEx.Text($"{ExcelItemHelper.GetName(x)} / 最低耐久度（不含兵装库）： {Utils.FindMinimumItemCondition(x, false)} / 最低耐久度（含兵装库）： {Utils.FindMinimumItemCondition(x, true)}");
                }
            }
        }
        if(ImGui.CollapsingHeader("受收藏柜保护的物品"))
        {
            foreach(var x in Utils.PlayerEntireInventory)
            {
                var cnt = InventoryManager.Instance()->GetInventoryContainer(x);
                List<uint> prot = [];
                for(var i = 0; i < cnt->Size; i++)
                {
                    var slot = cnt->GetInventorySlot(i);
                    if(slot->ItemId != 0)
                    {
                        if(S.CabinetManager.ShouldExcludeItemFromProcessing(slot->ItemId, slot->Condition))
                        {
                            prot.Add(slot->ItemId);
                        }
                    }
                }
                if(prot.Count > 0)
                {
                    if(ImGui.CollapsingHeader($"{x}##cab"))
                    {
                        ImGuiEx.Text(prot.Select(z => ExcelItemHelper.GetName(z, true)).Print("\n"));
                    }
                }
            }
        }
        if(ImGui.CollapsingHeader("受幻化台保护的物品"))
        {
            foreach(var x in Utils.PlayerEntireInventory)
            {
                var cnt = InventoryManager.Instance()->GetInventoryContainer(x);
                List<uint> prot = [];
                for(var i = 0; i < cnt->Size; i++)
                {
                    var slot = cnt->GetInventorySlot(i);
                    if(slot->ItemId != 0)
                    {
                        if(S.MirageManager.ShouldExcludeItemFromProcessing(slot->ItemId, slot->Condition))
                        {
                            prot.Add(slot->ItemId);
                        }
                    }
                }
                if(prot.Count > 0)
                {
                    if(ImGui.CollapsingHeader($"{x}##cab2"))
                    {
                        ImGuiEx.Text(prot.Select(z => ExcelItemHelper.GetName(z, true)).Print("\n"));
                    }
                }
            }
        }
    }
}
