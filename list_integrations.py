"""System report: list integrations with parent division, categories, and risk.
Format preserved: ID | Parent Division | Name | Vendor | Score | Need | Criticality | Installed | DR | Safety | Security | Monetary | Customer Service | Risk | Categories | Last Modified
"""
import database
database.initialize_database()

def fetch_integration_rows():
    conn = database.connect_db()
    cur = conn.cursor()
    # Join through integration_categories to categories; aggregate distinct category names
    cur.execute(
        """
     SELECT i.id,
         a.division AS parent_division,
         i.name,
         COALESCE(i.vendor, ''),
         i.score, i.need, i.criticality, i.installed,
         i.disaster_recovery, i.safety, i.security, i.monetary, i.customer_service,
         ROUND(COALESCE(i.risk_score, (10 - i.score) * i.criticality), 2) AS risk_score,
         COALESCE(GROUP_CONCAT(DISTINCT c.name), '') AS categories,
         i.last_modified
        FROM system_integrations i
        LEFT JOIN applications a ON i.parent_app_id = a.id
        LEFT JOIN integration_categories ic ON i.id = ic.integration_id
        LEFT JOIN categories c ON ic.category_id = c.id
        GROUP BY i.id
        ORDER BY LOWER(a.division), LOWER(i.name)
        """
    )
    rows = cur.fetchall()
    conn.close()
    return rows


def print_report():
    rows = fetch_integration_rows()
    header = [
        "ID", "Parent Division", "Name", "Vendor", "Score", "Need", "Criticality", "Installed", "DR",
        "Safety", "Security", "Monetary", "Customer Service", "Risk", "Categories", "Last Modified"
    ]
    print(" | ".join(header), flush=True)
    for r in rows:
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
        cat_str = _fmt(r[14])
        print(" | ".join([
            str(r[0]), r[1] or '', r[2], r[3], str(r[4]), str(r[5]), str(r[6]), str(r[7]), str(r[8]),
            str(r[9]), str(r[10]), str(r[11]), str(r[12]), str(r[13]), cat_str, str(r[15])
        ]), flush=True)

if __name__ == "__main__":
    print_report()
