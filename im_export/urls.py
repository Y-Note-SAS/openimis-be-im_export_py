from dateutil.relativedelta import relativedelta
from datetime import date
from datetime import timedelta

nb_periods, remainder = divmod(1000, 2500)
effective_date = date(2025, 6, 5)
print(f"there is the expiry date before {effective_date}")
# Calcul de la date d'expiration en fonction de la périodicité

base_expiry= effective_date + relativedelta(months=1)

grace_days = 2 * 30
grace_period = timedelta(days=grace_days) if grace_days else timedelta(0)
expiry_date = base_expiry + grace_period
print(f"there is the grace period {grace_period}")
print(f"there is the expiry date after {expiry_date}")
a = 0
if not a:
    print("Ok..")

