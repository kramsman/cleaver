""" test if we can identify new groups being added via pull_group """

# age = ['Old', 'Young']
# ethnicity = ['Black', 'Asian', 'Other']
# county = ['Kings', 'Queens', 'Nassau', 'Suffolk']
#
# county_header = ['{kings}','{queens}','{nassau}','{suffolk}',]
# county_header = ['{county}','{phone}','{url}',]

groups = [
    ["ethnicity", "age", "county", "pull_group", "addresses"],
    ["black", "young", "kings", 1, 234],
    ["black", "old", "kings", 1, 54],
    ["black", "old", "queens", 1, 112],
    ["asian", "young", "kings", 1, 75],

    ["black", "young", "kings", 2, 234],
    ["black", "old", "kings", 2, 54],
    ["black", "old", "queens", 2, 112],
    ["asian", "young", "kings", 2, 75],
    ["asian", "old", "kings", 2, 68],

    ["black", "young", "kings", 3, 234],
    ["black", "old", "kings", 3, 54],
    ["black", "old", "queens", 3, 112],
    ["asian", "young", "kings", 3, 75],
    ["asian", "old", "kings", 3, 21],
    ["asian", "old", "queens", 3, 99],
]

recent_pull_group = 2

groups = [ list(map(str, x)) for x in groups ]
factory_combos1_header = groups.pop(0)
factory_combos1 = [["-".join(tup[:-2])] + ["-".join(tup[:-3])] + list(tup) for tup in groups if int(tup[3]) <= recent_pull_group]

orig_groups = set([tup[0] for tup in factory_combos1 if int(tup[5]) < recent_pull_group])
recent_groups = set([tup[0] for tup in factory_combos1 if int(tup[5]) == recent_pull_group])

added_groups = recent_groups - orig_groups

a=1
