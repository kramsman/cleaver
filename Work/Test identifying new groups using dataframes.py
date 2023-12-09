""" test if we can identify new groups being added via pull_group """

import pandas
import pandas as pd

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
    ["black", "young", "nassau", 3, 54],
    ["black", "old", "kings", 3, 54],
    ["black", "old", "queens", 3, 112],
    ["asian", "young", "kings", 3, 75],
    ["asian", "old", "kings", 3, 21],
    ["asian", "old", "queens", 3, 99],
]

groups_header = groups.pop(0)
groups_df = pd.DataFrame(groups, columns=groups_header)

groupy_vars = ['county', 'ethnicity', 'age', ]
group_vars_remove = ['county',]
# groupy_vars = ['county']
# factory_vars = groupy_vars[:-1]  # removes last item
factory_vars = [x for x in groupy_vars if x not in group_vars_remove]

if groupy_vars:
    groups_df['campaign_var'] = groups_df[groupy_vars].agg('-'.join, axis=1)
else:
    groups_df['campaign_var'] = ""
if factory_vars:
    groups_df['factory_var'] = groups_df[factory_vars].agg('-'.join, axis=1)
else:
    groups_df['factory_var'] = ""

added_campaigns = groups_df[['campaign_var', 'pull_group']].drop_duplicates(subset=['campaign_var'], keep='first')
added_factory = groups_df[['factory_var', 'pull_group']].drop_duplicates(subset=['factory_var'], keep='first')
added_county = groups_df[['county', 'pull_group']].drop_duplicates(subset=['county'], keep='first')



xx= groups_df['group']
if isinstance(xx,pd.Series):
    xx = pd.DataFrame(xx)

recent_pull_group = 2

groups = [ list(map(str, x)) for x in groups ]
factory_combos1_header = groups.pop(0)
factory_combos1 = [["-".join(tup[:-2])] + ["-".join(tup[:-3])] + list(tup) for tup in groups if int(tup[3]) <= recent_pull_group]

orig_groups = set([tup[0] for tup in factory_combos1 if int(tup[5]) < recent_pull_group])
recent_groups = set([tup[0] for tup in factory_combos1 if int(tup[5]) == recent_pull_group])

added_groups = recent_groups - orig_groups

a=1
