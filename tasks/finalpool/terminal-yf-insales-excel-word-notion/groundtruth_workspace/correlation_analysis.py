# Эталонный correlation_analysis.py
GOLD_CHANGE_PCT = 27.01
OZON_CHANGE_PCT = 18.75
if GOLD_CHANGE_PCT > 50:
    print('Market conditions suggest CONTRACTING margins (gold inflation).')
else:
    print('Market conditions suggest MAINTAINING margins (limited cost pressure).')
