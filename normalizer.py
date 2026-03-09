def normalize(data):

    systems = []

    for s in data["systems"]:

        system = {}

        system["name"] = s["name"]
        system["module_count"] = len(s.get("modules",[]))
        system["interface_count"] = len(s.get("interfaces",[]))
        system["table_count"] = len(s.get("tables",[]))

        data_size = s.get("data_size_gb")

        if data_size is None:
            data_size = 50

        system["data_size_gb"] = data_size

        users = s.get("users",{})

        system["users"] = users.get("online_users",0)
        system["concurrent_users"] = users.get("concurrent_users",0)

        systems.append(system)

    return {"systems":systems}