from data_processor import aggregate_months

def build_dashboard(months):
    """Menghitung setiap bulan dengan rumus dashboard.js, lalu menjumlahkan komponen volume."""
    return aggregate_months(months)
