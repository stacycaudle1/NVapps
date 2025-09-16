"""System report: list applications with risk, business units, and categories.
Format preserved: ID | Division | Vendor | Score | Need | Criticality | Installed | DR | Safety | Security | Monetary | Customer Service | Risk | Business Units | Categories | Last Modified
"""
import database
database.initialize_database()

def fetch_application_rows():
    conn = database.connect_db()
    cur = conn.cursor()
    # Aggregate business units and categories
    cur.execute(
        """
     SELECT a.id,
         a.division,
         COALESCE(a.vendor, ''),
         a.score, a.need, a.criticality, a.installed,
         a.disaster_recovery, a.safety, a.security, a.monetary, a.customer_service,
         ROUND(COALESCE(a.risk_score, (10 - a.score) * a.criticality), 2) AS risk_score,
         COALESCE(GROUP_CONCAT(DISTINCT bu.name), '') AS business_units,
         COALESCE(GROUP_CONCAT(DISTINCT c.name), '') AS categories,
         a.last_modified
        FROM applications a
        LEFT JOIN application_business_units abu ON a.id = abu.app_id
        LEFT JOIN business_units bu ON abu.unit_id = bu.id
        LEFT JOIN application_categories ac ON a.id = ac.app_id
        LEFT JOIN categories c ON ac.category_id = c.id
        GROUP BY a.id
        ORDER BY LOWER(a.division)
        """
    )
    rows = cur.fetchall()
    conn.close()
    return rows


def print_report():
    rows = fetch_application_rows()
    header = [
        "ID", "Division", "Vendor", "Score", "Need", "Criticality", "Installed", "DR",
        "Safety", "Security", "Monetary", "Customer Service", "Risk", "Business Units", "Categories", "Last Modified"
    ]
    print(" | ".join(header), flush=True)
    for r in rows:
        # Normalize business units & categories to comma+space separated unique tokens preserving order
        def _fmt(csv_val):
            if not csv_val:
                return ''
            seen = set()
            ordered = []
            for part in str(csv_val).split(','):
                p = part.strip()
                if p and p.lower() not in seen:
                    seen.add(p.lower())
                    ordered.append(p)
            return ', '.join(ordered)
        bu_str = _fmt(r[13])
        cat_str = _fmt(r[14])
        print(" | ".join([
            str(r[0]), r[1], r[2], str(r[3]), str(r[4]), str(r[5]), str(r[6]), str(r[7]),
            str(r[8]), str(r[9]), str(r[10]), str(r[11]), str(r[12]), bu_str, cat_str, str(r[15])
        ]), flush=True)

if __name__ == "__main__":
    print_report()
