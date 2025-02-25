# %%
# Fetch historical data and go
# 
import pandas as pd
isin= pd.read_clipboard() 
# %%
securities= isin.copy()
securities= securities["RDIsin"].tolist()
securities= [sec + " Corp" for sec in securities] 


# %%
securities
# %%

# %%

from xbbg import blp
import blpapi
from blpapi import SessionOptions, Service
field_static= ["issuer","payment_rank","security_des", "issue_dt", "maturity", "mty_years_tdy", "asset_swap_spd_bid"]
df_description1= blp.bdp(securities, field_static)


# %% et la on y va 
# Vérifier les types de données
from asyncore import read
df= df_description1.copy()
df['issue_dt'] = pd.to_datetime(df['issue_dt'], errors='coerce')
df['maturity'] = pd.to_datetime(df['maturity'], errors='coerce')

df['mty_years_tdy'] = pd.to_numeric(df['mty_years_tdy'], errors='coerce')
df['asset_swap_spd_bid'] = pd.to_numeric(df['asset_swap_spd_bid'], errors= 'coerce')
print(df.dtypes)

# %%
# Étape 2: Créer des buckets de maturités
def create_maturity_buckets(df):
    bins = [0, 3, 5, 8, float('inf')]
    labels = ['<=3Y', '3-5Y', '5-8Y', '>8Y']
    df['maturity_bucket'] = pd.cut(df['mty_years_tdy'], bins=bins, labels=labels, right=False)
    return df
# %%
df = create_maturity_buckets(df)
grouped = df.groupby('issuer')


# step 3: we construct the pairs by issuer (Sr Preferred, Subordinated)
def create_pairs(group):
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

pairs = grouped.apply(create_pairs).explode().dropna().tolist()
print(pairs)
# %%


# %%
# step 4, here we compute the spread 
def calculate_spread(pair):
    sr_preferred, subordinated = pair
    spread = subordinated['asset_swap_spd_bid'].values[0] - sr_preferred['asset_swap_spd_bid'].values[0]
    return spread

spreads = [calculate_spread(pair) for pair in pairs]
# %%
# just for 
for i, pair in enumerate(pairs):
    sr_preferred, subordinated = pair
    print(f"Pair {i+1}:")
    print(f"Sr Preferred: {sr_preferred['issuer'].values[0]} - {sr_preferred['security_des'].values[0]}) - {sr_preferred['mty_years_tdy'].values[0]} years - {subordinated['ISIN'].values[0]}")

    print(f"Subordinated: {subordinated['issuer'].values[0]} - {subordinated['security_des'].values[0]}) - {subordinated['mty_years_tdy'].values[0]} years - {subordinated['ISIN'].values[0]}")
    print(f"Spread: {spreads[i]}")

# %%
for i, pair in enumerate(pairs):
    sr_preferred, subordinated = pair
    print(pair[0]["ISIN"].values[0])
# %%
# Vous pouvez également créer un DataFrame pour stocker les résultats
results = pd.DataFrame({
    'bucket': [pair[0]['maturity_bucket'].values[0] for pair in pairs],
    'ISIN-SNP': [pair[0]['ISIN'].values[0] for pair in pairs],
    'Sr Preferred Issuer': [pair[0]['issuer'].values[0] for pair in pairs],
    'Sr Preferred Maturity': [pair[0]['mty_years_tdy'].values[0] for pair in pairs],
    'Spread1': [pair[0]['asset_swap_spd_bid'].values[0] for pair in pairs],

    'ISIN-sub':[pair[1]['ISIN'].values[0] for pair in pairs],
    'Subordinated Issuer': [pair[1]['issuer'].values[0] for pair in pairs],
    'Subordinated Maturity': [pair[1]['mty_years_tdy'].values[0] for pair in pairs],
    'Spread2': [pair[1]['asset_swap_spd_bid'].values[0] for pair in pairs],
    'Spd2 - Spd1': spreads
    
})

 # %%
print(results)

# %%
results.to_clipboard()
# %%
