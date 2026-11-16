from __future__ import annotations

import json
import re
import sys
import argparse
import csv
from pathlib import Path
from typing import Any

from openpyxl import Workbook, load_workbook
from openpyxl.utils import get_column_letter


BASE_DIR = Path(__file__).resolve().parent
DATASET_PATH = BASE_DIR / "SYNTHETIC_DATASET.json"
EXCEL_OUTPUT_PATH = BASE_DIR / "filtered_leads_result.xlsx"

BLOCKED_CRM_STATUSES = {"跟进中", "已成交", "公海保护期内"}
SIGNAL_CATEGORIES = {
    "行业百强": ("百强", "50强"),
    "专精特新小巨人": ("专精特新", "小巨人"),
    "营收倍增/瞪羚企业": ("营收倍增", "瞪羚企业"),
    "纳税大户榜单": ("纳税大户", "纳税百强"),
    "荣誉资质": ("高新技术企业", "ISO"),
}
SUPPLIER_OR_CUSTOMER_FIELDS = ("is_supplier_of", "is_customer_of", "downstream_client_of")
PEER_FIELDS = ("competitor_of", "same_product_peer_of")
ANCHOR_LIST_FIELDS = ("same_list_with", "same_list_anchor")

COLUMN_ALIASES = {
    "company_name": "company_name",
    "公司名称": "company_name",
    "企业名称": "company_name",
    "crm_status": "crm_status",
    "crm跟进状态": "crm_status",
    "crm状态": "crm_status",
    "kp_name": "kp_name",
    "kp姓名": "kp_name",
    "联系人姓名": "kp_name",
    "决策人姓名": "kp_name",
    "kp_title": "kp_title",
    "kp职务": "kp_title",
    "职务": "kp_title",
    "contact_phone": "contact_phone",
    "kp_contact_phone": "contact_phone",
    "联系电话": "contact_phone",
    "kp电话": "contact_phone",
    "手机号": "contact_phone",
    "contact_email": "contact_email",
    "kp_email": "contact_email",
    "kp邮箱": "contact_email",
    "高管邮箱": "contact_email",
    "lists": "lists",
    "榜单": "lists",
    "榜单信号": "lists",
    "高价值信号": "lists",
    "tenders_last_6_months": "tenders_last_6_months",
    "近6个月招投标次数": "tenders_last_6_months",
    "近6个月招投标记录": "tenders_last_6_months",
    "is_supplier_of": "is_supplier_of",
    "relationships_is_supplier_of": "is_supplier_of",
    "上游供应商": "is_supplier_of",
    "供应商关系": "is_supplier_of",
    "is_customer_of": "is_customer_of",
    "downstream_client_of": "is_customer_of",
    "relationships_is_customer_of": "is_customer_of",
    "relationships_downstream_client_of": "is_customer_of",
    "下游客户": "is_customer_of",
    "客户关系": "is_customer_of",
    "competitor_of": "competitor_of",
    "same_product_peer_of": "competitor_of",
    "relationships_competitor_of": "competitor_of",
    "relationships_same_product_peer_of": "competitor_of",
    "同行": "competitor_of",
    "同行企业": "competitor_of",
    "same_list_with": "same_list_with",
    "same_list_anchor": "same_list_with",
    "relationships_same_list_with": "same_list_with",
    "同一榜单": "same_list_with",
    "同榜单客户": "same_list_with",
}


def _normalize_header(header: Any) -> str | None:
    if header is None:
        return None
    header_text = str(header).strip()
    if not header_text:
        return None
    if header_text in COLUMN_ALIASES:
        return COLUMN_ALIASES[header_text]

    year_match = re.search(r"20\d{2}", header_text)
    lower_header = header_text.lower()
    if year_match and ("revenue" in lower_header or "营收" in header_text):
        return f"revenue_{year_match.group(0)}"
    return header_text.lower()


def _optional_text(value: Any) -> str | None:
    if value is None:
        return None
    text = str(value).strip()
    return text or None


def _integer(value: Any) -> int:
    if isinstance(value, bool) or value in (None, ""):
        return 0
    if isinstance(value, int):
        return value
    match = re.search(r"-?\d+", str(value))
    return int(match.group(0)) if match else 0


def _signal_list(value: Any) -> list[str]:
    if isinstance(value, list):
        return [str(item).strip() for item in value if str(item).strip()]
    text = _optional_text(value)
    if not text:
        return []
    return [item.strip() for item in re.split(r"[;；|、,，]", text) if item.strip()]


def _row_is_empty(row: dict[Any, Any]) -> bool:
    return all(value in (None, "") for value in row.values())


def _read_csv(path: Path) -> list[dict[str, Any]]:
    with path.open("r", encoding="utf-8-sig", newline="") as file:
        rows = [dict(row) for row in csv.DictReader(file) if not _row_is_empty(dict(row))]
    return rows


def _read_xlsx(path: Path) -> list[dict[str, Any]]:
    workbook = load_workbook(path, read_only=True, data_only=True)
    sheet = workbook.active
    rows_iter = sheet.iter_rows(values_only=True)
    headers = next(rows_iter, None)
    if headers is None:
        workbook.close()
        return []

    rows = [
        dict(zip(headers, values))
        for values in rows_iter
        if values and not all(value in (None, "") for value in values)
    ]
    workbook.close()
    return rows


def _company_from_row(raw_row: dict[str, Any]) -> dict[str, Any]:
    row: dict[str, str] = {}
    for header, value in raw_row.items():
        normalized_header = _normalize_header(header)
        if normalized_header and value not in (None, ""):
            row.setdefault(normalized_header, value)

    financials = {
        header: row[header]
        for header in row
        if re.fullmatch(r"revenue_20\d{2}", header)
    }
    relationships = {}
    for field in SUPPLIER_OR_CUSTOMER_FIELDS:
        if row.get(field):
            relationships[field] = _optional_text(row[field])
    for field in PEER_FIELDS:
        if row.get(field):
            relationships[field] = _optional_text(row[field])
    for field in ANCHOR_LIST_FIELDS:
        if row.get(field):
            relationships[field] = _optional_text(row[field])

    return {
        "company_name": _optional_text(row.get("company_name")),
        "financials": financials,
        "kp": {
            "name": _optional_text(row.get("kp_name")),
            "title": _optional_text(row.get("kp_title")),
            "contact_phone": _optional_text(row.get("contact_phone")),
            "contact_email": _optional_text(row.get("contact_email")),
        },
        "lists": _signal_list(row.get("lists")),
        "tenders_last_6_months": _integer(row.get("tenders_last_6_months")),
        "relationships": relationships,
        "crm_status": _optional_text(row.get("crm_status")),
    }


def _load_companies(path: Path) -> list[dict[str, Any]]:
    suffix = path.suffix.lower()
    if suffix == ".json":
        return _load_json(path)
    if suffix == ".csv":
        raw_rows = _read_csv(path)
    elif suffix == ".xlsx":
        raw_rows = _read_xlsx(path)
    else:
        raise ValueError("输入文件仅支持 .json、.csv 或 .xlsx")

    companies = [_company_from_row(row) for row in raw_rows if _optional_text(row.get("company_name"))]
    if not companies:
        raise ValueError("输入文件中没有可识别的公司记录")
    return companies


def _load_json(path: Path) -> Any:
    with path.open("r", encoding="utf-8") as file:
        return json.load(file)


def _revenue_series(financials: dict[str, Any]) -> list[tuple[int, float]]:
    series: list[tuple[int, float]] = []
    for key, value in financials.items():
        match = re.fullmatch(r"revenue_(\d{4})", key)
        if not match or value is None:
            continue
        revenue = float(value)
        if revenue > 0:
            series.append((int(match.group(1)), revenue))
    return sorted(series)


def _annual_declines(series: list[tuple[int, float]]) -> list[tuple[int, float]]:
    declines: list[tuple[int, float]] = []
    for (previous_year, previous_revenue), (year, revenue) in zip(series, series[1:]):
        if previous_revenue > 0:
            decline_rate = (previous_revenue - revenue) / previous_revenue * 100
            declines.append((year, decline_rate))
    return declines


def _cagr(series: list[tuple[int, float]]) -> float | None:
    if len(series) < 2:
        return None
    first_year, first_revenue = series[0]
    last_year, last_revenue = series[-1]
    if last_year <= first_year or first_revenue <= 0 or last_revenue <= 0:
        return None
    years = last_year - first_year
    return (last_revenue / first_revenue) ** (1 / years) - 1


def _rejection_reason(company: dict[str, Any]) -> str | None:
    crm_status = str(company.get("crm_status", "")).strip()
    if crm_status in BLOCKED_CRM_STATUSES:
        return f"CRM状态为“{crm_status}”，存在跟进冲突"

    declines = _annual_declines(_revenue_series(company.get("financials", {})))
    for year, decline_rate in declines:
        if decline_rate > 20:
            return f"{year}年营收同比下滑{decline_rate:.1f}%，超过20%风险阈值"

    key_person = company.get("kp") or {}
    if not key_person.get("name") or not key_person.get("title"):
        return "关键人线索为空"

    return None


def _matched_signal_categories(lists: list[str]) -> list[str]:
    categories: list[str] = []
    for category, keywords in SIGNAL_CATEGORIES.items():
        if any(keyword in item for item in lists for keyword in keywords):
            categories.append(category)
    return categories


def _financial_score_details(series: list[tuple[int, float]]) -> tuple[int, dict[str, Any] | None]:
    cagr = _cagr(series)
    if cagr is None:
        return 0, None

    if cagr > 0.20:
        points = 25
    elif 0.10 <= cagr <= 0.20:
        points = 15
    elif 0 <= cagr < 0.10:
        points = 5
    else:
        return 0, None

    first_year, first_revenue = series[0]
    last_year, last_revenue = series[-1]
    evidence = (
        f"CAGR为{cagr * 100:.1f}%；{first_year}年营收{first_revenue:g}，"
        f"{last_year}年营收{last_revenue:g}；来源: SYNTHETIC_DATASET.json financials"
    )
    return points, {
        "rule_name": "财务健康度",
        "points": points,
        "evidence": evidence,
    }


def _signal_score_details(company: dict[str, Any]) -> tuple[int, dict[str, Any] | None]:
    lists = [str(item) for item in company.get("lists", []) if str(item).strip()]
    categories = _matched_signal_categories(lists)
    tenders = int(company.get("tenders_last_6_months", 0) or 0)
    signal_count = len(categories) + (1 if tenders > 0 else 0)
    if signal_count == 0:
        return 0, None

    points = min(40, signal_count * 10)
    evidence_parts: list[str] = []
    if categories:
        evidence_parts.append(f"命中榜单类别: {'、'.join(categories)}；原始榜单: {'、'.join(lists)}")
    if tenders > 0:
        evidence_parts.append(f"近6个月招投标记录{tenders}次")
    evidence_parts.append("来源: SYNTHETIC_DATASET.json lists/tenders_last_6_months")

    return points, {
        "rule_name": "榜单与高价值信号",
        "points": points,
        "evidence": "；".join(evidence_parts),
    }


def _relationship_score_details(company: dict[str, Any]) -> tuple[int, dict[str, Any] | None]:
    relationships = company.get("relationships") or {}
    supplier_or_customer = next(
        (relationships[field] for field in SUPPLIER_OR_CUSTOMER_FIELDS if relationships.get(field)),
        None,
    )
    peer = next((relationships[field] for field in PEER_FIELDS if relationships.get(field)), None)
    anchor = next((relationships[field] for field in ANCHOR_LIST_FIELDS if relationships.get(field)), None)

    if supplier_or_customer:
        return 25, {
            "rule_name": "关系网络裂变",
            "points": 25,
            "evidence": f"与已合作客户“{supplier_or_customer}”存在供应链上下游关系；来源: SYNTHETIC_DATASET.json relationships",
        }
    if peer:
        return 15, {
            "rule_name": "关系网络裂变",
            "points": 15,
            "evidence": f"数据集标记为已合作客户“{peer}”的同行；来源: SYNTHETIC_DATASET.json relationships",
        }
    if anchor:
        return 10, {
            "rule_name": "关系网络裂变",
            "points": 10,
            "evidence": f"与已合作客户“{anchor}”命中同一榜单；来源: SYNTHETIC_DATASET.json relationships",
        }
    return 0, None


def _key_person_score_details(company: dict[str, Any]) -> tuple[int, dict[str, Any] | None]:
    key_person = company.get("kp") or {}
    phone = str(key_person.get("contact_phone") or "").strip()
    email = str(key_person.get("contact_email") or key_person.get("email") or "").strip()
    phone_digits = re.sub(r"\D", "", phone)
    is_direct_mobile = len(phone_digits) == 11 and phone_digits.startswith("1")
    if not is_direct_mobile and not email:
        return 0, None

    contact_evidence = "数据源包含KP直连手机号" if is_direct_mobile else "数据源包含高管邮箱"
    return 10, {
        "rule_name": "关键人可触达性",
        "points": 10,
        "evidence": f"{contact_evidence}；来源: SYNTHETIC_DATASET.json kp",
    }


def _score_company(company: dict[str, Any]) -> tuple[int, list[dict[str, Any]]]:
    score_details: list[dict[str, Any]] = []
    financial_detail = _financial_score_details(_revenue_series(company.get("financials", {})))
    signal_detail = _signal_score_details(company)
    relationship_detail = _relationship_score_details(company)
    key_person_detail = _key_person_score_details(company)
    for detail in (financial_detail, signal_detail, relationship_detail, key_person_detail):
        if detail[1] is not None:
            score_details.append(detail[1])

    total_score = sum(detail["points"] for detail in score_details)
    return total_score, score_details


def _rejected_result(company: dict[str, Any], reason: str) -> dict[str, Any]:
    key_person = company.get("kp") or {}
    return {
        "company_name": company.get("company_name", ""),
        "total_score": 0,
        "is_rejected": True,
        "reject_reason": reason,
        "score_details": [],
        "kp_info": {
            "name": key_person.get("name", ""),
            "title": key_person.get("title", ""),
            "contact_phone": key_person.get("contact_phone"),
            "contact_email": key_person.get("contact_email"),
        },
    }


def _scored_result(company: dict[str, Any], total_score: int, score_details: list[dict[str, Any]]) -> dict[str, Any]:
    if total_score < 40:
        return _rejected_result(company, f"总分{total_score}低于40分淘汰线")

    final_level = "S级" if total_score >= 80 else "A级" if total_score >= 60 else "B级"
    key_person = company.get("kp") or {}
    return {
        "company_name": company.get("company_name", ""),
        "final_level": final_level,
        "total_score": total_score,
        "is_rejected": False,
        "reject_reason": "",
        "score_details": score_details,
        "kp_info": {
            "name": key_person.get("name", ""),
            "title": key_person.get("title", ""),
            "contact_phone": key_person.get("contact_phone"),
            "contact_email": key_person.get("contact_email"),
        },
    }


def screen_prospects(companies: list[dict[str, Any]]) -> list[dict[str, Any]]:
    results: list[dict[str, Any]] = []
    for company in companies:
        rejection_reason = _rejection_reason(company)
        if rejection_reason:
            results.append(_rejected_result(company, rejection_reason))
            continue
        total_score, score_details = _score_company(company)
        results.append(_scored_result(company, total_score, score_details))

    return sorted(results, key=lambda item: (item["is_rejected"], -item["total_score"]))


def _validate_output(report: dict[str, Any]) -> None:
    if not isinstance(report, dict):
        raise TypeError("顶层输出必须是对象")
    if set(("total_count", "prospects")) - report.keys():
        raise ValueError("顶层输出缺少 total_count 或 prospects")
    if not isinstance(report["total_count"], int) or isinstance(report["total_count"], bool):
        raise TypeError("total_count必须是整数")
    if not isinstance(report["prospects"], list):
        raise TypeError("prospects必须是数组")
    if report["total_count"] != len(report["prospects"]):
        raise ValueError("total_count必须等于prospects数量")

    required_fields = {"company_name", "total_score", "is_rejected"}
    for result in report["prospects"]:
        missing_fields = required_fields - result.keys()
        if missing_fields:
            raise ValueError(f"输出缺少必填字段: {', '.join(sorted(missing_fields))}")
        if not isinstance(result["company_name"], str):
            raise TypeError("company_name必须是字符串")
        if not isinstance(result["total_score"], int):
            raise TypeError("total_score必须是整数")
        if not isinstance(result["is_rejected"], bool):
            raise TypeError("is_rejected必须是布尔值")


def _write_excel_report(results: list[dict[str, Any]], path: Path) -> None:
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "筛选结果"
    headers = [
        "公司名",
        "最终评级",
        "得分",
        "一票否决原因",
        "主要加分项",
        "KP姓名",
        "KP职务",
        "KP联系电话",
        "KP邮箱",
    ]
    sheet.append(headers)

    for result in results:
        score_summary = "；".join(
            f"{detail['rule_name']}+{detail['points']}分（{detail['evidence']}）"
            for detail in result.get("score_details", [])
        )
        key_person = result.get("kp_info") or {}
        sheet.append([
            result.get("company_name", ""),
            "淘汰" if result.get("is_rejected") else result.get("final_level", ""),
            result.get("total_score", 0),
            result.get("reject_reason", ""),
            score_summary,
            key_person.get("name", ""),
            key_person.get("title", ""),
            key_person.get("contact_phone") or "",
            key_person.get("contact_email") or "",
        ])

    column_widths = [34, 10, 8, 30, 80, 14, 16, 18, 24]
    for index, width in enumerate(column_widths, start=1):
        sheet.column_dimensions[get_column_letter(index)].width = width
    sheet.freeze_panes = "A2"
    workbook.save(path)


def _parse_arguments() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="B2B销售客户智能筛选原型")
    parser.add_argument(
        "--input",
        type=Path,
        default=DATASET_PATH,
        help="输入文件路径，支持 .json / .csv / .xlsx",
    )
    parser.add_argument(
        "--excel-output",
        type=Path,
        default=EXCEL_OUTPUT_PATH,
        help="Excel报告输出路径，默认 filtered_leads_result.xlsx",
    )
    return parser.parse_args()


def main() -> None:
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8")

    arguments = _parse_arguments()
    companies = _load_companies(arguments.input)
    results = screen_prospects(companies)
    report = {"total_count": len(results), "prospects": results}
    _validate_output(report)
    _write_excel_report(results, arguments.excel_output)
    print(json.dumps(report, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
