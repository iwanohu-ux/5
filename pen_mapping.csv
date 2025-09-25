// Archicad DevKit / C++17
#include "APIEnvir.h"
#include "ACAPinc.h"

#include <array>
#include <vector>
#include <string>
#include <fstream>
#include <sstream>
#include <unordered_map>
#include <cctype>

// ------- ユーティリティ（トリム・CSVパース） -------
static inline std::string LTrim(std::string s){ s.erase(s.begin(), std::find_if(s.begin(), s.end(), [](unsigned char c){return !std::isspace(c);})); return s; }
static inline std::string RTrim(std::string s){ s.erase(std::find_if(s.rbegin(), s.rend(), [](unsigned char c){return !std::isspace(c);}).base(), s.end()); return s; }
static inline std::string Trim(std::string s){ return RTrim(LTrim(std::move(s))); }

static std::vector<std::string> SplitCSVLine(const std::string& line)
{
    std::vector<std::string> cells;
    std::string cur; bool inQ=false;
    for (size_t i=0;i<line.size();++i){
        char c=line[i];
        if (inQ){
            if (c=='"'){
                if (i+1<line.size() && line[i+1]=='"'){ cur.push_back('"'); ++i; } // "" -> "
                else inQ=false;
            } else cur.push_back(c);
        } else {
            if (c=='"') inQ=true;
            else if (c==','){ cells.push_back(cur); cur.clear(); }
            else cur.push_back(c);
        }
    }
    cells.push_back(cur);
    for (auto& t: cells) t = Trim(t);
    return cells;
}

// BOM除去（UTF-8想定）
static void StripUTF8BOM(std::string& s){
    if (s.size()>=3 && (unsigned char)s[0]==0xEF && (unsigned char)s[1]==0xBB && (unsigned char)s[2]==0xBF) {
        s.erase(0,3);
    }
}

// ヘッダ名を探す（日本語／英語どちらでも可）
static bool ResolveHeaderCols(const std::vector<std::string>& header, int& idxOld, int& idxNew)
{
    auto norm = [](std::string x){
        for (auto& c: x) { c = (char)std::tolower((unsigned char)c); }
        // 全角矢印やスペース混入の簡易吸収
        // 例: "→" カラムが紛れ込んでも無視できるように
        return Trim(x);
    };
    std::unordered_map<std::string,int> m;
    for (int i=0;i<(int)header.size();++i) m[norm(header[i])] = i;

    auto findCol = [&](std::initializer_list<const char*> names, int& out)->bool{
        for (auto* n: names){
            auto it = m.find(norm(n));
            if (it!=m.end()){ out=it->second; return true; }
        }
        return false;
    };

    // 日本語 / 英語どちらでもOK
    bool ok1 = findCol({"旧index","oldindex","旧ｉｎｄｅｘ"}, idxOld);
    bool ok2 = findCol({"新index","newindex","新ｉｎｄｅｘ"}, idxNew);
    return ok1 && ok2;
}

// 文字列→short 変換（安全側）
static bool ToShort(const std::string& s, short& out)
{
    if (s.empty()) return false;
    char* end=nullptr;
    long v = std::strtol(s.c_str(), &end, 10);
    if (end==s.c_str() || v<0 || v>32767) return false;
    out = static_cast<short>(v);
    return true;
}

// ------- CSV読み込み→ペン置換マップ生成 -------
static bool BuildPenMapFromCSV(const std::string& csvPath, std::array<short,257>& outMap, GS::UniString* errMsg=nullptr)
{
    // 既定は恒等マップ
    for (short i=0;i<=256;++i) outMap[i]=i;

    std::ifstream ifs(csvPath, std::ios::in | std::ios::binary);
    if (!ifs){
        if (errMsg) *errMsg = GS::UniString("CSVを開けません: ") + GS::UniString(csvPath.c_str());
        return false;
    }

    std::string line;
    if (!std::getline(ifs, line)){ if (errMsg) *errMsg = "CSVが空です"; return false; }
    StripUTF8BOM(line);
    auto header = SplitCSVLine(line);

    int idxOld=-1, idxNew=-1;
    if (!ResolveHeaderCols(header, idxOld, idxNew)){
        if (errMsg) *errMsg = "ヘッダに「旧Index / 新Index」が見つかりません（OldIndex/NewIndex も可）";
        return false;
    }

    size_t row=1;
    while (std::getline(ifs, line)){
        ++row;
        if (Trim(line).empty()) continue;
        auto cells = SplitCSVLine(line);
        if ((int)cells.size() <= std::max(idxOld, idxNew)) continue;

        short oldIdx=0, newIdx=0;
        if (!ToShort(cells[idxOld], oldIdx) || !ToShort(cells[idxNew], newIdx)) {
            // 無効行はスキップ（報告はログへ）
            ACAPI_WriteReport("pen_mapping.csv: 無効な行 #%zu （旧Index/新Index を数値にできません）", false, (GS::UInt32)row);
            continue;
        }
        if (oldIdx<1 || oldIdx>256 || newIdx<1 || newIdx>256) {
            ACAPI_WriteReport("pen_mapping.csv: 範囲外のペン番号（1..256） row=%zu old=%d new=%d", false, (GS::UInt32)row, oldIdx, newIdx);
            continue;
        }
        outMap[oldIdx] = newIdx;
    }
    return true;
}

// ------- 要素ペンのリマップ（前回提示の骨格を簡略転記） -------
static inline void RemapPen(short& p, const std::array<short,257>& map){ if (p>=1 && p<=256) p = map[p]; }

static void RemapCommonPens(API_Element& el, const std::array<short,257>& map)
{
    RemapPen(el.header.pen,         map);
    RemapPen(el.header.contPen,     map);
    RemapPen(el.header.sectContPen, map);
}

static void RemapTypeSpecificPens(API_Element& el, API_ElementMemo* /*memo*/, const std::array<short,257>& map)
{
    switch (el.header.typeID) {
        case API_LineID:        RemapPen(el.line.linePen, map); break;
        case API_ArcID:         RemapPen(el.arc.linePen, map); break;
        case API_PolyLineID:    RemapPen(el.polyLine.linePen, map); break;
        case API_HatchID:       RemapPen(el.hatch.linePen, map); RemapPen(el.hatch.fillPen, map); RemapPen(el.hatch.bkgPen, map); break;
        case API_TextID:        RemapPen(el.text.color.pen, map); break;
        case API_LabelID:       RemapPen(el.label.pen, map); break;
        case API_WallID:        RemapPen(el.wall.contPen, map); RemapPen(el.wall.corePen, map); RemapPen(el.wall.fillerPen, map); break;
        case API_BeamID:        RemapPen(el.beam.contPen, map); RemapPen(el.beam.corePen, map); break;
        case API_ColumnID:      RemapPen(el.column.contPen, map); RemapPen(el.column.corePen, map); break;
        case API_SlabID:        RemapPen(el.slab.contPen, map); RemapPen(el.slab.corePen, map); break;
        case API_RoofID:        RemapPen(el.roof.contPen, map); break;
        case API_MeshID:        RemapPen(el.mesh.contPen, map); break;
        case API_DimensionID:   RemapPen(el.dimension.pen, map); break;
        default: break; // 階段/手摺/カーテンウォール/ゾーンなど必要に応じ追加
    }
}

static void RemapOneElement(const API_Guid& guid, const std::array<short,257>& map)
{
    API_Element el = {}; el.header.guid = guid;
    if (ACAPI_Element_Get(&el) != NoError) return;

    API_ElementMemo memo = {}; GS::AutoPtr<API_ElementMemo> memoGuard(&memo);
    (void)ACAPI_Element_GetMemo(guid, &memo, APIMemoMask_All);

    RemapCommonPens(el, map);
    RemapTypeSpecificPens(el, &memo, map);

    API_Element mask = {};
    ACAPI_ELEMENT_MASK_SET(mask, API_ElemHeaderType, pen);
    ACAPI_ELEMENT_MASK_SET(mask, API_ElemHeaderType, contPen);
    ACAPI_ELEMENT_MASK_SET(mask, API_ElemHeaderType, sectContPen);

    (void)ACAPI_Element_Change(&el, &mask, &memo, 0, true);
}

// ------- エントリ：CSVを指定して全要素リマップ -------
GSErrCode DoRemapAllPensFromCSV(const std::string& csvPath)
{
    std::array<short,257> map{};
    GS::UniString err;
    if (!BuildPenMapFromCSV(csvPath, map, &err)){
        ACAPI_WriteReport("ペンマップ作成に失敗: %T", true, &err);
        return APIERR_GENERAL;
    }

    GS::Array<API_Guid> guids;
    GSErrCode e = ACAPI_Element_GetElemList(API_ZombieElemID, &guids);
    if (e != NoError) return e;

    return ACAPI_CallUndoableCommand("Remap all pens (CSV)", [&]() -> GSErrCode {
        for (const auto& g : guids) RemapOneElement(g, map);
        return NoError;
    });
}
