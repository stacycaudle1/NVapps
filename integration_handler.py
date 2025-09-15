def handle_integration(cur, conn, app_id, app_name, int_name, int_vendor, int_score, int_need, int_criticality, 
                    int_installed, int_dr, int_safety, int_security, int_monetary, int_customer_service, last_mod, rownum):
    """Helper function to handle integration creation with proper error handling."""
    from datetime import datetime
    
    if not int_name or not app_id:
        return False
        
    try:
        # Calculate risk score as average of numeric fields
        numeric_fields = [
            int_score or 0, int_need or 0, int_criticality or 0,
            int_installed or 0, int_dr or 0, int_safety or 0,
            int_security or 0, int_monetary or 0, int_customer_service or 0
        ]
        risk_score = sum(numeric_fields) / len(numeric_fields) if numeric_fields else 0
        
        # Insert with all fields explicitly specified
        cur.execute('''
            INSERT INTO system_integrations 
            (parent_app_id, name, vendor, score, need, criticality,
             installed, disaster_recovery, safety, security,
             monetary, customer_service, notes, risk_score, last_modified)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        ''', (
            app_id, int_name, int_vendor,
            int_score or 0, int_need or 0, int_criticality or 0,
            int_installed or 0, int_dr or 0, int_safety or 0,
            int_security or 0, int_monetary or 0, int_customer_service or 0,
            '', risk_score, last_mod or datetime.now().isoformat()
        ))
        conn.commit()
        print(f"DEBUG: Successfully created integration {int_name} for app {app_name}")
        return True
        
    except Exception as e:
        print(f"DEBUG: Failed to create integration {int_name}: {e}")
        return False