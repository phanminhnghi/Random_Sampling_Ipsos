import pandas as pd
import numpy as np

def random_allocate_samples(file_path, city_sample_dict, min_per_ward, output_file, ward_priority_ratio):
    df = pd.read_excel(file_path, sheet_name="3321-2025")
    df.columns = df.columns.str.strip()

    result_rows = []

    for city_code, total_samples in city_sample_dict.items():
        city_df = df[df['Mã TP'] == city_code]
        
        phuong_wards = city_df[city_df['Cấp'] == "Phường"]['Tên'].dropna().tolist()
        xa_wards = city_df[city_df['Cấp'] == "Xã"]['Tên'].dropna().tolist()

        phuong_ratio = ward_priority_ratio[city_code]["Phường"]

        wards_needed = total_samples // min_per_ward
        phuong_needed = int(wards_needed * phuong_ratio)
        xa_needed = wards_needed - phuong_needed

        if phuong_needed > len(phuong_wards) or xa_needed > len(xa_wards):
            raise ValueError(f"Không đủ phường/xã để phân bổ cho {city_df['Tỉnh / Thành Phố'].iloc[0]}!")

        selected_phuong = np.random.choice(phuong_wards, size=phuong_needed, replace=False) if phuong_needed > 0 else []
        selected_xa = np.random.choice(xa_wards, size=xa_needed, replace=False) if xa_needed > 0 else []
        selected_wards = np.concatenate([selected_phuong, selected_xa])

        allocation = {ward: min_per_ward for ward in selected_wards}
        samples_allocated = len(selected_wards) * min_per_ward
        samples_left = total_samples - samples_allocated

        while samples_left > 0:
            ward = np.random.choice(selected_wards)
            allocation[ward] += 1
            samples_left -= 1

        for ward, count in allocation.items():
            ward_info = city_df[city_df['Tên'] == ward].iloc[0]
            result_rows.append({
                'Tỉnh / Thành Phố': ward_info['Tỉnh / Thành Phố'] if 'Tỉnh / Thành Phố' in ward_info else ward_info['Tên'],
                'Mã TP': city_code,
                'Phường/Xã': ward,
                'Số lượng mẫu': count
            })

    result_df = pd.DataFrame(result_rows)
    result_df.to_excel(output_file, index=False)
    return result_df

#_____________________________________
#_______ input information ___________
#_____________________________________
Input_city_sample_dict = {
    79: 3000,
    1: 80
}

Input_min_per_ward = 10

Input_ward_priority_ratio = {
    79: {"Phường": 0.8, "Xã": 0.2},
    1: {"Phường": 0.9, "Xã": 0.1}
}
#_____________________________________

random_allocate_samples(
    file_path='Source/TinhThanhVietNam2025.xlsx',
    city_sample_dict=Input_city_sample_dict,
    min_per_ward=Input_min_per_ward,
    output_file='Result_Sampling.xlsx',
    ward_priority_ratio=Input_ward_priority_ratio
)