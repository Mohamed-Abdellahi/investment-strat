# %%
# Fetch historical data and go
# 
isin= pd.read_clipboard() 
# %%
securities= isin.copy()
securities= securities["RDIsin"].tolist()
securities= [sec + " Corp" for sec in securities] 

# %%

# %%

# %%

from distutils.log import error
from xbbg import blp
import blpapi
from blpapi import SessionOptions, Service
field_static= ["issuer","ticker", "cntry_of_risk","payment_rank","security_des", "issue_dt", "maturity", "mty_years_tdy", "asset_swap_spd_bid", "workout_date_mid_to_worst"]
df_description1= blp.bdp(securities, field_static)


# %% et la on y va 
# Vérifier les types de données
from asyncore import read
import pandas as pd
#df= df_description1.copy()
df=pd.read_clipboard()
df['issue_dt'] = pd.to_datetime(df['issue_dt'], errors='coerce')
df['maturity'] = pd.to_datetime(df['maturity'], errors='coerce')
df["workout_date_mid_to_worst"]= pd.to_datetime(df["workout_date_mid_to_worst"], errors= 'coerce')

df['workout_dt_years_mid'] = pd.to_numeric(df['workout_dt_years_mid'], errors='coerce')

df['asset_swap_spd_bid'] = pd.to_numeric(df['asset_swap_spd_bid'], errors= 'coerce')
print(df.dtypes)


# %% 
# Étape 2: Créer des buckets de maturités
#def create_maturity_buckets(df):
#    bins = [0, 3, 5, 8, float('inf')]  ## this will determine the inteervals of our cut
   # labels = ['<=3Y', '3-5Y', '5-8Y', '>8Y']
   # df['maturity_bucket'] = pd.cut(df[workout_dt_years_mid], bins=bins, #labels=labels, right=False)
   # return df
# %%
#df = create_maturity_buckets(df)
#grouped = df.groupby('issuer')


# step 3: we construct the pairs by issuer (Sr Preferred, Subordinated)
def create_pairs_by_buckets(group):
    pairs = []
    for bucket in group['maturity_bucket'].unique():
        print(f"Processing bucket: {bucket}")
        bucket_group = group[group['maturity_bucket'] == bucket]
        print(f"Bucket group: {bucket_group}")
        sr_preferred = bucket_group[bucket_group['payment_rank'] == 'Sr Preferred']
        subordinated = bucket_group[bucket_group['payment_rank'] == 'Subordinated']
        print(f"Sr Preferred: {sr_preferred}")
        print(f"Subordinated: {subordinated}")
        if not sr_preferred.empty and not subordinated.empty:
            pairs.append((sr_preferred, subordinated))
    return pairs

def create_pairs_by_country(df, mty_years_tdy_limit= 0.25, constraint1= 'cntry_of_risk'):
    constraint1= 'cntry_of_risk'
    constraint2= 'ticker'
    constraint3= '' ## cap structure
    pairs = []
    sr_preferred = df[df['payment_rank'] == 'Sr Preferred']
    subordinated = df[df['payment_rank'] == 'Subordinated']

    for _, sr_row in sr_preferred.iterrows():
        for _, sub_row in subordinated.iterrows():
            if (abs(sr_row['workout_dt_years_mid'] - sub_row['workout_dt_years_mid']) <= mty_years_tdy_limit) and (sr_row[constraint1] == sub_row[constraint1]):
                pairs.append((sr_row, sub_row))

    return pairs

# %%

def create_pairs_by_constr(df, workout_dt_years_limit=0.25, constraints=None):
    if constraints is None:
        constraints = []

    pairs = []
    sr_preferred = df[df['payment_rank'] == 'Sr Preferred']
    subordinated = df[df['payment_rank'] == 'Subordinated']

    for _, sr_row in sr_preferred.iterrows():
        for _, sub_row in subordinated.iterrows():
            if abs(sr_row['workout_dt_years_mid'] - sub_row['workout_dt_years_mid']) <= workout_dt_years_limit:
                constraints_met = True
                for col, values in constraints:
                    if values is None:
                        continue  # Ignorer cette contrainte si values est None
                    if isinstance(values, list):
                        if sr_row[col] not in values or sub_row[col] not in values:
                            constraints_met = False
                            break
                    else:
                        if sr_row[col] != values or sub_row[col] != values:
                            constraints_met = False
                            break
                if constraints_met:
                    pairs.append((sr_row, sub_row))

    return pairs

# %%
def create_pairs(group): ## group is a dataframe
    pairs= []

    sr_preferred= group[group['payment_rank'] =='Sr Preferred']
    subordinated= group[group['payment_rank'] == 'Subordinated']

    for _, sp_row in sr_preferred.iterrows():
        for _, sub_row in subordinated.iterrows():
          if abs(sp_row['workout_dt_years_mid'] - sub_row['workout_dt_years_mid']) <= 0.25:
                pairs.append((sp_row, sub_row))

    return pairs  


#pairs = grouped.apply(create_pairs).explode().dropna().tolist()



# step 4, here we compute the spread 
def calculate_spread(pair):
    sr_preferred, subordinated = pair
    spread = subordinated['asset_swap_spd_bid'] - sr_preferred['asset_swap_spd_bid']
    return spread


# %%
def create_results_dataframe(pairs, spreads):
    results = pd.DataFrame({
        'Country-SNP': [pair[0]['cntry_of_risk'] for pair in pairs],
        'ISIN-SNP': [pair[0]['ISIN'] for pair in pairs],
        'Sr Preferred Issuer': [pair[0]['issuer'] for pair in pairs],
        'Sr Preferred Maturity': [pair[0]['workout_dt_years_mid'] for pair in pairs],
        'Spread1': [pair[0]['asset_swap_spd_bid'] for pair in pairs],

        'Country-Sub': [pair[1]['cntry_of_risk'] for pair in pairs],
        'ISIN-sub': [pair[1]['ISIN'] for pair in pairs],
        'Subordinated Issuer': [pair[1]['issuer'] for pair in pairs],
        'Subordinated Maturity': [pair[1]['workout_dt_years_mid'] for pair in pairs],
        'Spread2': [pair[1]['asset_swap_spd_bid'] for pair in pairs],
        'Spd2 - Spd1': spreads
    })
    return results
# %%



constraints = [
    # ('cntry_of_risk', ['FR', 'IT']),  # La colonne 'cntry_of_risk' doit être 'IT' ou 'FR'
    ('TICKER',"AIB"),                # La colonne 'ticker' doit être 'TICK1'
   # ('issuer', None)                    # Ignorer la contrainte sur 'issuer'
]


pairs = create_pairs_by_constr(df, workout_dt_years_limit=0.25, constraints=constraints)
spreads = [calculate_spread(pair) for pair in pairs]
results_df = create_results_dataframe(pairs, spreads)
# put it in a dataframes

# %%
results_df.to_clipboard()
 # %%
results_df

# %%

# %%
