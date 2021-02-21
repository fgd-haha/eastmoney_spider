import json


def save_to_json(output_dir, sorted_dict):
    with open(f"{output_dir}/sell_result.json", "w", encoding="utf8") as f:
        temp_dict = {}
        for code in sorted_dict[0:20]:
            temp_dict[code[0]] = (code[1]["ShareSZ_Chg_One"], code[1]["SName"])
            # temp_dict[code[0]] = code[1]
        json.dump(temp_dict, f, ensure_ascii=False, indent=2)
    with open(f"{output_dir}/buy_result.json", "w", encoding="utf8") as f:
        temp_dict = {}
        for code in sorted_dict[-1:-21:-1]:
            temp_dict[code[0]] = (code[1]["ShareSZ_Chg_One"], code[1]["SName"])
            # temp_dict[code[0]] = code[1]
        json.dump(temp_dict, f, ensure_ascii=False, indent=2)
