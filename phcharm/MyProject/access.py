access_list = scenario2.GetExistingAccesses() #获取所有的Access
for access_path in tqdm(access_list):
    # access_path:('Satellite/Sat0_0/Transmitter/Transmitter_Sat0_0','Satellite/Sat0_1/Receiver/Reciver_Sat0_1',True)
    # access_path[0]:发射机地址
    # access_path[1]：接收机地址
    Transmitter_name = access_path[0].split('/')[-1]
    Reciver_name = access_path[1].split('/')[-1]
    access = scenario2.GetAccessBetweenObjectsByPath(access_path[0], access_path[1])
    Transmitter_name = access_path[0].split('/')[-1]
    Reciver_name = access_path[1].split('/')[-1]
    access = scenario2.GetAccessBetweenObjectsByPath(access_path[0], access_path[1])
