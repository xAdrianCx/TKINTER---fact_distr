from configparser import ConfigParser 

cfg = ConfigParser()

cfg_path = "C:\\Users\\User\\Desktop\\New folder\\config\\cfg.ini"
cfg.read(cfg_path)
# tariff_c1 = cfg.get("TARIFF", "tariff_c1")
# print(tariff_c1)
# tariff_c2 = cfg.get("TARIFF", "tariff_c2")
# print(tariff_c2)
# tariff_c3 = cfg.get("TARIFF", "tariff_c3")
# print(tariff_c3)

sup = dict(cfg.items("SUPPLIERS"))
for key, value in sup.items():
	if key == "energy distribution services".lower():
		print(key, value)
