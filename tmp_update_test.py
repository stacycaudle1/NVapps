import database

integration_id = 6
fields = {
    'name': 'SmokeIntegration',
    'vendor': 'SmokeVendor',
    'score': 5,
    'need': 4,
    'criticality': 10,
    'installed': 2,
    'disaster_recovery': 1,
    'safety': 2,
    'security': 3,
    'monetary': 4,
    'customer_service': 5,
    'notes': 'Updated from test'
}

try:
    res = database.update_system_integration(integration_id, fields)
    print('update_system_integration returned:', res)
except Exception as e:
    import traceback
    traceback.print_exc()
    print('Exception:', e)
