import bca_function as bf


def run_bca():
    # 根据各种数据计算BCA浓度，并且新建一个Excel，将结果写入
    bf.notice_excel()
    bf.load_original_excel()
    bf.dilution_ratio()
    bf.sample_volume()
    bf.write_excel()
    bf.calculate_bca()


run_bca()

# Todo 继续重构代码
# Todo 构建UI界面并关联代码
# Todo 与RT-PCR项目融合
