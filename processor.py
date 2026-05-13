import pandas as pd
import numpy as np
import random
import io
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from datetime import date


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def fix_date(col):
    def conv(v):
        if pd.isna(v):
            return pd.NaT
        if isinstance(v, (int, float)):
            return pd.Timestamp("1899-12-30") + pd.Timedelta(days=float(v))
        return pd.Timestamp(v)
    return col.apply(conv)


def ordinal(n):
    """Return ordinal string for an integer, e.g. 1 -> '1st', 11 -> '11th', 22 -> '22nd'."""
    if 11 <= (n % 100) <= 13:
        suffix = "th"
    else:
        suffix = {1: "st", 2: "nd", 3: "rd"}.get(n % 10, "th")
    return f"{n}{suffix}"

def format_date_range(dates):
    """Return a human-readable date range string like 'May 1st thru May 3rd 2026'."""
    sorted_dates = sorted(dates)
    def fmt(d):
        return f"{d.strftime('%B')} {ordinal(d.day)}"
    if len(sorted_dates) == 1:
        return f"{fmt(sorted_dates[0])} {sorted_dates[0].strftime('%Y')}"
    return f"{fmt(sorted_dates[0])} thru {fmt(sorted_dates[-1])} {sorted_dates[-1].strftime('%Y')}"


VENDOR_REPLACEMENTS = [
    ("AXS", "Veritix"), ("AXS.com", "Veritix"),
    ("Ticketmaster.com", "Ticketmaster"), ("Toyota Center- Houston", "Veritix"),
    ("Ak-Chin Pavilion", "Live Nation Ak-Chin Pavilion"),
    ("BB&T Pavilion", "Live Nation BB&T Pavilion"),
    ("Dos Equis Pavilion", "Live Nation Dos Equis Pavilion"),
    ("Jiffy Lube Live", "Live Nation Jiffy Lube Live"),
    ("Jacobs Pavilion", "Live Nation Jacobs Pavilion"),
    ("MIDFLORIDA Credit Union Amp", "Live Nation MidFlorida Credit Union Amp"),
    ("Shoreline Amphitheatre", "Live Nation Shoreline Amphitheatre"),
    ("Ticketmaster Phones", "Ticketmaster"),
    ("Shubert Organization Telecharge", "Telecharge"),
    ("Veritix.com", "Veritix"),
    ("Huntington Bank Pavilion", "Live Nation Huntington Bank"),
    ("Darien Lake Amphitheatre", "Live Nation Darien Lake Amphitheatre"),
    ("Leader Bank Pavillion", "Live Nation Leader Bank Pavilion"),
    ("Coastal Credit Union Music Park at Walnut Creek", "Live Nation MidFlorida Credit Union Amp"),
    ("TD Pavilion at the Mann", "Live Nation TD Pavilion at the Mann"),
    ("Live Nation", "Ticketmaster"),
    ("Xfinity Theatre", "Live Nation Xfinity Theatre CT"),
    ("Live Nation Xfinity Boston", "Live Nation Xfinity Center Boston"),
    ("Live Nation PNC Bank", "Live Nation PNC Bank Charlotte"),
    ("LN Ruoff Home Mortgage Music Center", "Live Nation Ruoff"),
    ("Alpine Valley", "Live Nation Alpine Valley"),
    ("The Pavilion at Toyota Music Factory", "Live Nation Pavilion at Toyota Music Factory"),
    ("Live Nation Pnc Charlotte", "Live Nation PNC Music Pavilion"),
    ("The Pavilion at Star Lake", "Live Nation Pavilion At Star Lake"),
    ("Bank of New Hampshire Pavilion", "Live Nation Bank of New Hampshire"),
    ("North Island Credit Union", "Live Nation North Island Credit Union"),
    ("Northwell Health at Jones Beach Theater", "Live Nation Jones Beach"),
    ("PNC Music Pavilion", "Live Nation PNC Music Pavilion"),
    ("Ruoff Music Center", "Live Nation Ruoff Music Center"),
    ("Waterfront Music Pavilion", "Live Nation BB&T Pavilion"),
    ("Concord Pavilion", "Live Nation Concord Pavilion"),
    ("Bethel Woods", "Live Nation Bethel Woods"),
    ("Ford Idaho Center Ampitheater", "Live Nation Ford Idaho Amp"),
    ("The Met Philadelphia", "Live Nation The Met Philadelphia"),
    ("The Masonic", "Live Nation Masonic"),
    ("FPL Solar Ampitheater at Bayfront Park", "Live Nation FPL Solar Amp"),
    ("Xfinity Center", "Live Nation Xfinity Center Boston"),
    ("The Cynthia Woods Mitchell Pavilion", "Live Nation Cynthia Woods Mitchell Pavilion"),
    ("Toyota Amphitheatre", "Live Nation Toyota Amp"),
    ("Cellairis Amphitheatre at Lakewood", "Live Nation Cellairis"),
    ("Oak Mountian Amphitheatre", "Live Nation Oak Mountain"),
    ("Chicago Fire", "Chicago Fire FC"), ("Dallas FC", "FC Dallas"),
    ("Houston Dynamo", "Houston Dynamo FC"), ("LAFC", "Los Angeles FC"),
    ("LA Galaxy", "Los Angeles Galaxy"), ("Minnesota United", "Minnesota United FC"),
    ("New York FC", "New York City FC"), ("Seattle Sounders", "Seattle Sounders FC"),
    ("St. Louis City", "St. Louis City SC"), ("Vancouver Whitecaps", "Vancouver Whitecaps FC"),
    ("Portland Trailblazers", "Portland Trail Blazers"),
    ("Concert Extras", "Live Nation Extras"),
    ("PHILADELPHIA 76ERS", "Philadelphia 76ers"),
    ("SAN FRANCISCO 49RS", "San Francisco 49ers"),
    ("Concert Partials", "Concert Seasons"), ("Legacy StubHub", "StubHub"),
]

BROADWAY_VENUES = {
    "Brooks Atkinson Theatre": "Box Office - Brooks Atkinson Theatre",
    "Hudson Broadway Theatre": "Box Office - Hudson Broadway Theatre",
    "Minskoff Theatre": "Box Office - Minskoff Theatre",
    "Madison Square Garden": "Box Office - MSG Advance",
    "Neil Simon Theatre": "Box Office - Neil Simon Theatre",
    "Richard Rodgers Theatre": "Box Office - Richard Rodgers Theatre",
    "Winter Garden Theater New York": "Box Office - Winter Garden Theatre",
    "Walter Kerr Theatre": "Box Office - Walter Kerr Theatre",
    "August Wilson Theatre": "Box Office - August Wilson Theatre",
    "Music Box Theatre New York": "Box Office - Music Box Theatre",
    "Imperial Theatre New York": "Box Office - Imperial Theatre",
    "Lyceum Theatre New York": "Box Office - Lyceum Theatre",
    "Booth Theatre": "Box Office - Booth Theatre",
    "Longacre Theatre": "Box Office - Longacre Theatre",
    "Majestic Theatre New York": "Box Office - Majestic Theatre",
    "Broadhurst Theatre": "Box Office - Broadhurst Theatre",
    "Stephen Sondheim Theatre": "Box Office - Stephen Sondheim Theatre",
    "Lena Horne Theatre": "Box Office - Lena Horne Theatre",
    "Lyric Theatre - NY": "Box Office - Lyric Theatre",
    "Winter Garden Theatre (Toronto)": "Box Office - Winter Garden Theatre (Toronto)",
    "Marquis Theatre New York": "Box Office - Marquis Theatre",
    "Lunt Fontanne Theatre": "Box Office - Lunt Fontanne Theatre",
    "John Golden Theatre": "Box Office - John Golden Theatre",
    "Circle In The Square": "Box Office - Circle In The Square",
}

CONCERT_SEASONS_MAP = {
    "Ak-Chin Pavilion": "Live Nation Ak-Chin Pavilion",
    "Alpine Valley Music Theatre": "Live Nation Alpine Valley",
    "Bank of New Hampshire Pavilion": "Live Nation Bank of New Hampshire",
    "Darling's Waterfront Pavilion": "Live Nation Waterfront",
    "Darlings Waterfront Pavilion": "Live Nation Waterfront",
    "Bethel Woods Center For The Arts": "Live Nation Bethel Woods",
    "Blossom Music Center": "Live Nation Blossom MC",
    "Lakewood Amphitheatre": "Live Nation Lakewood Amphitheatre",
    "Coastal Credit Union Music Park at Walnut Creek": "Live Nation Coastal Credit Union",
    "Concord Pavilion": "Live Nation Concord Pavilion",
    "Cynthia Woods Mitchell Pavilion": "Live Nation Cynthia Woods Mitchell Pavilion",
    "Darien Lake Amphitheater": "Live Nation Darien Lake Amphitheatre",
    "Dos Equis Pavilion": "Live Nation Dos Equis Pavilion",
    "FivePoint Amphitheatre": "Live Nation FivePoint",
    "Ford Idaho Center": "Live Nation Ford Idaho Amp",
    "Bayfront Park Amphitheatre": "Live Nation FPL Solar Amp",
    "Glen Helen Amphitheater": "Live Nation Glen Helen",
    "Gorge Amphitheatre": "Live Nation Gorge",
    "Hershey Theatre": "Live Nation Hershey",
    "Hollywood Casino Amphitheatre - Tinley Park": "Live Nation Hollywood Casino - Tinley Park",
    "Huntington Bank Pavilion at Northerly Island": "Live Nation Huntington Bank",
    "Isleta Amphitheater": "Live Nation Isleta Amphitheatre",
    "Jacobs Pavilion at Nautica": "Live Nation Jacobs Pavilion",
    "Jiffy Lube Live": "Live Nation Jiffy Lube Live",
    "Northwell Health at Jones Beach Theater": "Live Nation Jones Beach",
    "KeyBank Center": "Live Nation KeyBank",
    "Leader Bank Pavilion": "Live Nation Leader Bank Pavilion",
    "The Masonic": "Live Nation Masonic",
    "MGM Music Hall at Fenway": "Live Nation MGM Music Hall",
    "MidFlorida Credit Union Amphitheatre": "Live Nation MidFlorida Credit Union Amp",
    "North Island Credit Union Amphitheatre": "Live Nation North Island Credit Union",
    "Oak Mountain Amphitheatre": "Live Nation Oak Mountain",
    "The Pavilion at Star Lake": "Live Nation Pavilion At Star Lake",
    "Pavilion at the Toyota Music Factory": "Live Nation Pavilion At Toyota Music Factory",
    "PNC Bank Arts Center": "Live Nation PNC Bank Charlotte",
    "PNC Music Pavilion": "Live Nation PNC Music Pavilion",
    "Ruoff Music Center": "Live Nation Ruoff",
    "Shoreline Amphitheatre": "Live Nation Shoreline Amphitheatre",
    "TD Pavilion at the Mann": "Live Nation TD Pavilion at the Mann",
    "The Met Philadelphia": "Live Nation The Met Philadelphia",
    "Toyota Amphitheatre": "Live Nation Toyota Amp",
    "White River Amphitheatre": "Live Nation White River",
    "The Wiltern": "Live Nation Wiltern",
    "Xfinity Center": "Live Nation Xfinity Center Boston",
    "Xfinity Theatre": "Live Nation Xfinity Theatre CT",
    "Freedom Mortgage Pavilion": "Live Nation Freedom Mortgage Music Pavilion",
    "Bayfront Park-Miami": "Live Nation FPL Solar Amp",
    "Idaho Center Amphitheater": "Live Nation Ford Idaho Amp",
    "Coca-Cola Roxy": "Live Nation Coca-Cola Roxy Theatre",
    "Farm Bureau Insurance Lawn at White River State Park": "Live Nation TCU Amp",
    "Charlotte Metro Credit Union Amphitheatre": "Live Nation Charlotte Metro Credit Union",
    "713 Music Hall": "Live Nation 713 Music Hall",
    "TCU Amphitheater at White River State Park": "Live Nation TCU Amp",
    "USANA Amphitheatre": "Live Nation USANA Amp",
    "Hollywood Casino Amphitheater - St. Louis": "Live Nation Hollywood Casino - St. Louis",
    "Hollywood Casino Amphitheatre St Louis": "Live Nation Hollywood Casino - St. Louis",
    "iTHINK Financial Amphitheatre": "Live Nation iTHINK Financial Amp",
    "Ameris Bank Amphitheatre": "Live Nation Ameris Bank Amp",
    "Ascend Amphitheater": "Live Nation Ascend Amp",
    "FirstBank Amphitheater": "Live Nation FirstBank Amp",
    "Hersheypark Stadium": "Live Nation Hersheypark Stadium",
    "Veterans United Home Loans Amphitheater": "Live Nation Veterans United Amp",
    "Starlight Theatre": "Live Nation Starlight Theatre",
    "Riverbend Music Center": "Live Nation Riverbend Music Center",
    "The Terminal - Houston": "Live Nation 713 Music Hall",
    "Veterans United Home Loans Amphitheater at Virginia Beach": "Live Nation Veterans United Amp",
    "Saint Louis Music Park": "Live Nation Saint Louis Music Park",
    "Old National Centre": "Live Nation Old National Centre",
    "Skyla Credit Union": "Live Nation Skyla Credit Union Amp",
    "Skyla Credit Union Amphitheatre": "Live Nation Skyla Credit Union Amp",
    "Pine Knob Music Center": "Live Nation Pine Knob Music Center",
    "St. Joseph's Health Amphitheater at Lakeview": "Live Nation St. Joseph's Health Amp",
    "Red Hat Amphitheater": "Live Nation Red Hat Amphitheater",
    "Talking Stick Resort Amphitheatre": "Live Nation Talking Stick Resort Amp",
    "The Fillmore Detroit": "Live Nation The Fillmore Detroit",
    "CFG Bank Arena": "Live Nation CFG Bank Arena",
    "Aragon Ballroom": "Live Nation Aragon Ballroom",
    "Saratoga Performing Arts Center": "Live Nation Saratoga Springs PAC",
    "The Fillmore - Charlotte": "Live Nation The Fillmore Charlotte",
    "Fillmore Auditorium-CO": "Live Nation The Fillmore Denver",
    "The Fillmore - Philadelphia": "Live Nation The Fillmore Philly",
    "The Fillmore-Silver Spring": "Live Nation The Fillmore Silver Spring",
    "SAP Center at San Jose": "Live Nation SAP Center at San Jose",
    "Arizona Federal Theatre": "Live Nation Arizona Federal Theatre",
    "Flagstar at Westbury Music Fair": "Live Nation Flagstar At Westbury Music Fair",
    "Uptown Minneapolis": "Live Nation Uptown Minneapolis",
    "The Pavilion At Toyota Music Factory": "Live Nation Pavilion At Toyota Music Factory",
    "Forest Hills Stadium": "Forest Hills Stadium",
    "NRG Stadium": "Houston Rodeo",
    "Hayden Homes Amphitheater": "Live Nation Hayden Homes Amphitheater",
    "The Dome at Oakdale Theatre": "Live Nation Toyota Oakdale Theatre",
    "The Dome at Toyota Oakdale Theatre": "Live Nation Toyota Oakdale Theatre",
    "Broadview Stage at SPAC": "Live Nation Broadview Stage at SPAC",
    "BECU Live Outdoor Venue": "Live Nation BECU",
    "BankNH Pavilion": "Live Nation Bank of New Hampshire",
    "Everwise Amphitheater at White River State Park": "Live Nation White River",
    "The Cynthia Woods Mitchell Pavilion presented by Huntsman": "Live Nation Cynthia Woods Mitchell Pavilion",
    "Harbor Yard Amphitheater": "Live Nation Harbor Yard Amp",
    "Toyota Pavilion at Concord": "Live Nation Concord Pavilion",
    "Utah First Credit Union Amphitheatre (formerly USANA Amp)": "Live Nation Usana Amp",
    "Toyota Oakdale Theatre": "Live Nation Toyota Oakdale Theatre",
    "Byline Bank Aragon Ballroom": "Live Nation Aragon Ballroom",
    "Skyline Stage at the Mann": "Live Nation Skyline Stage At The Mann",
    "Santa Barbara Bowl": "Live Nation Santa Barbara Bowl",
    "20 Monroe Live": "Live Nation GLC Live at 20 Monroe",
    "MIDFLORIDA Credit Union Amphitheatre at the FL State Fairgrounds": "Live Nation MidFlorida Credit Union Amp",
    "Credit Union 1 Amphitheatre": "Live Nation Credit Union 1 Amphitheatre",
    "Daily's Place": "Live Nation Dailys Place",
    "Greek Theatre Los Angeles": "Live Nation Greek Theatre Los Angeles",
    "Koka Booth Field 2": "Live Nation Koka Booth Amphitheatre",
    "Michigan Lottery Amphitheatre at Freedom Hill": "Live Nation Michigan Lottery Amphitheatre",
    "Mountain Winery": "Live Nation Mountain Winery",
    "Skyla Credit Union Amphitheatre at AvidXchange Music Factory": "Live Nation Skyla Credit Union Amp",
    "Vibrant Music Hall": "Live Nation Vibrant Music Hall",
    "Vina Robles Amphitheatre": "Live Nation Vina Robles Amphitheatre",
    "Whitewater Amphitheater": "Live Nation Whitewater Amphitheater",
    "Old National Centre Complex": "Live Nation Old National Centre",
    "Fiddlers Green Amphitheatre": "Live Nation Fiddlers Green Amphitheatre",
    "Pine Knob Music Theatre": "Live Nation Pine Knob Music Center",
    "Hollywood Palladium": "Live Nation Hollywood Palladium",
    "Hartford HealthCare Amphitheater": "Live Nation Hartford Healthcare Amphitheater (Harbor Yard)",
    "Mystic Lake Amphitheater": "Live Nation Mystic Lake Amphitheater",
    "Fillmore Minneapolis": "Live Nation Fillmore Minneapolis",
    "Uptown Theater - Minneapolis": "Live Nation Uptown Minneapolis",
    "Old National Centre.": "Live Nation Old National Centre",
}

BROADWAY_SEASONS_MAP = {
    "Boston Opera House": "Broadway Boston",
    "Colonial Theatre Boston": "Broadway Boston",
    "Fox Theatre - Atlanta": "Broadway Atlanta",
    "Hippodrome Theatre": "Broadway Baltimore",
    "BJCC Concert Hall": "Broadway Birmingham",
    "Shea's Buffalo Theatre": "Broadway Buffalo",
    "Procter and Gamble Hall at Aronoff Center for the Arts": "Broadway Cincinnati",
    "AT&T Performing Arts Center - Winspear Opera House": "Broadway Dallas",
    "Music Hall at Fair Park": "Broadway Dallas",
    "Durham Performing Arts Center": "Broadway Durham",
    "Broward Center Amaturo": "Broadway Ft Lauderdale",
    "Devos Performance Hall": "Broadway Grand Rapids",
    "Hollywood Pantages Theatre": "Broadway Hollywood",
    "Music Hall - Kansas City": "Broadway Kansas City",
    "Muriel Kauffman Theatre at Kauffman Center for the Performing Arts": "Broadway Kansas City",
    "Saenger Theatre-New Orleans": "Broadway New Orleans",
    "Sarofim Hall at The Hobby Center": "Broadway Houston",
    "Uihlein Hall at Marcus Center for the Performing Arts": "Broadway Milwaukee",
    "Orpheum Theatre Minneapolis": "Broadway Minneapolis",
    "Clowes Memorial Hall": "Broadway Indianapolis",
    "Old National Centre": "Broadway Indianapolis",
    "Paramount Theatre": "Broadway Seattle",
    "San Diego Civic Theatre": "Broadway San Diego",
    "San Jose Center for the Performing Arts": "Broadway San Jose",
}

MLB_TEAMS = {
    "Arizona Diamondbacks", "Atlanta Braves", "Baltimore Orioles", "Boston Red Sox",
    "Chicago White Sox", "Chicago Cubs", "Cincinnati Reds", "Cleveland Indians",
    "Colorado Rockies", "Detroit Tigers", "Houston Astros", "Kansas City Royals",
    "Los Angeles Angels", "Los Angeles Dodgers", "Miami Marlins", "Milwaukee Brewers",
    "Minnesota Twins", "New York Yankees", "New York Mets", "Oakland Athletics",
    "Philadelphia Phillies", "Pittsburgh Pirates", "San Diego Padres", "San Francisco Giants",
    "Seattle Mariners", "St. Louis Cardinals", "Tampa Bay Rays", "Texas Rangers",
    "Toronto Blue Jays", "Washington Nationals",
}

# Company → (sheet name, list of Company values in data)
COMPANY_SHEETS = {
    "Y&S":         ["YS Tickets", "YS-Seatgeek", "YS Tickets Spec", "YS-Seatgeek2"],
    "Grossman":    ["YSM Tickets"],
    "Sternbuch":   ["YSS Tickets"],
    "Pollak":      ["Pollak Tickets"],
    "Levine":      ["Yoni Levine"],
    "Levovitz":    ["Levovitz"],
    "GK":          ["GK LLC"],
    "Ticket Guy":  ["The Ticket Guy", "The Ticket Guy-Jas", "The Ticket Guy-Legacy", "The Ticket Guy VIP"],
    "Chase":       ["Jacks YS"],
    "YSA":         ["YSA", "YSA 2", "YSA 3"],
    "Katz":        ["YS Katz"],
    "TL":          ["YS TL"],
    "Waxler":      ["YSW"],
    "Damona":      ["Damon and Crew"],
}


# ---------------------------------------------------------------------------
# Excel styling
# ---------------------------------------------------------------------------

def write_sheet(wb, name, dataframe):
    ws = wb.create_sheet(name)
    cols = list(dataframe.columns)
    thin = Side(style="thin", color="CCCCCC")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    header_font = Font(name="Arial", bold=True, color="FFFFFF", size=10)
    header_fill = PatternFill("solid", start_color="4472C4")
    header_align = Alignment(horizontal="center", vertical="center", wrap_text=True)

    for ci, col in enumerate(cols, 1):
        cell = ws.cell(row=1, column=ci, value=col)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = header_align
        cell.border = border

    fill_odd  = PatternFill("solid", start_color="FFFFFF")
    fill_even = PatternFill("solid", start_color="EEF2FF")

    for ri, row in enumerate(dataframe.itertuples(index=False), 2):
        row_fill = fill_even if ri % 2 == 0 else fill_odd
        for ci, val in enumerate(row, 1):
            col_name = cols[ci - 1]
            if col_name == "PO Created" and val is not None and not (isinstance(val, float) and np.isnan(val)):
                try:
                    val = pd.Timestamp(val).strftime("%m/%d/%Y")
                except Exception:
                    pass
            cell = ws.cell(row=ri, column=ci, value=val)
            cell.font = Font(name="Arial", size=10)
            cell.alignment = Alignment(vertical="center")
            cell.border = border
            cell.fill = row_fill
            if col_name == "Total Cost":
                cell.number_format = "0.00"

    for ci, col in enumerate(cols, 1):
        max_len = len(str(col))
        for row in dataframe.itertuples(index=False):
            val = row[ci - 1]
            max_len = max(max_len, len(str(val)) if val is not None else 0)
        ws.column_dimensions[get_column_letter(ci)].width = min(max_len + 2, 55)

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions


# ---------------------------------------------------------------------------
# Core transformation
# ---------------------------------------------------------------------------

def apply_vendor_replacements(df):
    df["Vendor"] = df["Vendor"].replace("Box Office", "Default Vendor")
    for old, new in VENDOR_REPLACEMENTS:
        df["Vendor"] = df["Vendor"].str.replace(old, new, regex=False)
    return df


def build_all_query(df_raw):
    df = df_raw.copy()
    df["PO Created"] = fix_date(df["PO Created"]).dt.normalize().dt.date

    df["Ext PO #"] = df["Ext PO #"].fillna(" ").astype(str)
    df["PO Email Account"] = df["PO Email Account"].fillna(" ").astype(str)
    df = df.drop(columns=["Notes"], errors="ignore")
    df["Notes"] = (df["Team/Performer"].astype(str) + " / " +
                   df["PO Email Account"].astype(str) + " /  " +
                   df["Ext PO #"].astype(str))
    df["Notes (Short)"] = df["PO Email Account"].astype(str) + " / " + df["Ext PO #"].astype(str)

    drop_cols = ["PO #", "Opponent/Performer", "Event Date", "Sec", "Row", "Seats",
                 "Qty", "Cost", "Cancelled", "Created", "User", "Ext PO #", "PO Email Account"]
    df = df.drop(columns=[c for c in drop_cols if c in df.columns])

    df["Account"] = "Inventory Asset"
    df["Memo"] = df["Team/Performer"]
    df = df[["Company", "PO Created", "Account", "Vendor", "Team/Performer",
              "Memo", "Total Cost", "Notes", "Notes (Short)", "Delivery Type", "Venue", "Tags"]]

    df = apply_vendor_replacements(df)

    for old, new in [("Toyota Amphitheatre", "Live Nation Toyota Amp"),
                     ("Xfinity Center", "Live Nation Xfinity Center Boston"),
                     ("FPL Solar Ampitheater at Bayfront Park", "Live Nation FPL Solar Amp")]:
        df["Company"] = df["Company"].replace(old, new)
    df["Company"] = df["Company"].astype(str)
    df["Team/Performer"] = df["Team/Performer"].str.replace("Miami HEAT", "Miami Heat", regex=False)

    # Sports Extras → venue
    df["Vendor"] = np.where(df["Vendor"] == "Sports Extras", df["Venue"], df["Vendor"])

    # Ticket Guy broadway box office
    def ticket_guy_vendor(row):
        if row["Company"] != "The Ticket Guy":
            return row["Vendor"]
        if row["Vendor"] != "Default Vendor":
            return row["Vendor"]
        if "New World Stages" in str(row["Venue"]):
            return "Box Office - New World Stages"
        return BROADWAY_VENUES.get(row["Venue"], row["Vendor"])
    df["Vendor"] = df.apply(ticket_guy_vendor, axis=1)

    # Concert Seasons
    df["Vendor"] = df.apply(
        lambda r: CONCERT_SEASONS_MAP.get(r["Venue"], r["Vendor"]) if r["Vendor"] == "Concert Seasons" else r["Vendor"], axis=1)

    # Broadway Seasons — collapse Team/Performer, Notes, Memo to "Various" so all rows group into one bill
    is_bs = df["Vendor"] == "Broadway Seasons"
    df.loc[is_bs, "Team/Performer"] = "Various"
    df.loc[is_bs, "Memo"] = "Various"
    df.loc[is_bs, "Notes"] = "Various"
    df.loc[is_bs, "Notes (Short)"] = "Various"
    df["Vendor"] = df.apply(
        lambda r: BROADWAY_SEASONS_MAP.get(r["Venue"], r["Vendor"]) if r["Vendor"] == "Broadway Seasons" else r["Vendor"], axis=1)

    # MLB / Tickets.com
    df["MLB?"] = df["Team/Performer"].apply(lambda x: "Yes" if x in MLB_TEAMS else "No")
    df["Vendor"] = df.apply(
        lambda r: (r["Venue"] if r["MLB?"] == "No" else r["Team/Performer"]) if r["Vendor"] == "Tickets.com" else r["Vendor"], axis=1)
    df = df.drop(columns=["Venue", "MLB?"])

    # Proper case
    df["Vendor"] = df["Vendor"].str.title()
    df["Vendor"] = df["Vendor"].str.replace("Philadelphia 76Ers", "Philadelphia 76ers", regex=False)
    df["Vendor"] = df["Vendor"].str.replace("San Francisco 49Ers", "San Francisco 49ers", regex=False)

    # First groupby
    group_keys = ["Company", "PO Created", "Account", "Vendor", "Team/Performer",
                  "Memo", "Notes", "Notes (Short)"]
    df = df.groupby(group_keys, as_index=False, dropna=False)["Total Cost"].sum()
    df = df[df["Total Cost"] > 0]
    df["Vendor"] = df["Vendor"].astype(str)

    # Ticketmaster CAD
    df["VendorNew"] = df.apply(
        lambda r: ("Ticketmaster CAD" if str(r["Notes"]).endswith(("/TOR", "/VAN", "/QUE")) else "Ticketmaster")
        if r["Vendor"] == "Ticketmaster" else r["Vendor"], axis=1)
    df = df.rename(columns={"Notes": "Notes (Final)"})
    df = df.drop(columns=["Notes (Short)", "Vendor"])
    df = df.rename(columns={"VendorNew": "Vendor"})

    # Notes (Final) becomes Team/Performer (per M code)
    df = df[["Company", "PO Created", "Account", "Vendor", "Notes (Final)", "Total Cost"]]
    df = df.rename(columns={"Notes (Final)": "Team/Performer"})

    # Seasons tagging
    df["Seasons"] = df["Vendor"].apply(
        lambda v: "LN Extras" if v == "Live Nation Extras" else ("Live Nation" if "Live Nation" in str(v) else ""))
    df["Broadway_tag"] = df["Vendor"].apply(
        lambda v: "Broadway Extras" if v == "Broadway Extras" else ("Broadway" if "Broadway" in str(v) else ""))
    df["Seasons"] = df["Seasons"] + df["Broadway_tag"]
    df = df.drop(columns=["Broadway_tag"])

    # Collapse LN seasons rows
    df["Team/Performer"] = df.apply(
        lambda r: "Various / Various" if r["Seasons"] in ("Live Nation", "LN Extras") else r["Team/Performer"], axis=1)
    df = df.rename(columns={"Team/Performer": "Memo2"})

    # Second groupby
    group_keys2 = ["Company", "PO Created", "Account", "Vendor", "Memo2", "Seasons"]
    df = df.groupby(group_keys2, as_index=False, dropna=False)["Total Cost"].sum()

    df["Team/Performer"] = df["Memo2"] + " (" + df["Company"] + ")"
    df["Memo"] = df["Team/Performer"]
    df["Bill No."] = [random.randint(10000000, 99999999) for _ in range(len(df))]

    return df[["Company", "Bill No.", "PO Created", "Account", "Vendor",
               "Memo2", "Team/Performer", "Memo", "Total Cost", "Seasons"]]


def build_summary_query(df_raw):
    s = df_raw.copy()
    s["PO Created"] = fix_date(s["PO Created"]).dt.normalize().dt.date
    s["Account"] = "Inventory Asset"
    s = apply_vendor_replacements(s)
    s["Vendor"] = s["Vendor"].str.replace("FrontGate Tickets", "Front Gate Tickets", regex=False)
    s["MLB?"] = s["Team/Performer"].apply(lambda x: "Yes" if x in MLB_TEAMS else "No")
    s["Vendor"] = s.apply(
        lambda r: (r["Venue"] if r["MLB?"] == "No" else r["Team/Performer"]) if r["Vendor"] == "Tickets.com" else r["Vendor"], axis=1)

    grp = s.groupby(["Company", "PO Created", "Account", "Vendor", "Team/Performer"],
                    as_index=False, dropna=False).agg(Total_Cost=("Total Cost", "sum"), Qty=("Qty", "sum"))
    grp = grp.rename(columns={"Total_Cost": "Total Cost"})
    grp["Total Cost"] = grp["Total Cost"].round(2)
    grp = grp[["Company", "PO Created", "Account", "Vendor", "Team/Performer", "Qty", "Total Cost"]]
    return grp.sort_values(["Company", "PO Created", "Vendor", "Team/Performer"])


def filter_company(all_df, companies, rename_company=None, vendor_replace=None, strip_company_prefix=None):
    f = all_df[all_df["Company"].isin(companies)].copy()
    if rename_company:
        for old, new in rename_company.items():
            f["Company"] = f["Company"].str.replace(old, new, regex=False)
    if vendor_replace:
        for old, new in vendor_replace.items():
            f["Vendor"] = f["Vendor"].str.replace(old, new, regex=False)
    if strip_company_prefix:
        f["Company"] = f["Company"].str.replace(strip_company_prefix, "", regex=False)
    return f


# ---------------------------------------------------------------------------
# Main entry point
# ---------------------------------------------------------------------------

def process_file(file_bytes):
    """
    Takes raw bytes of the uploaded .xlsm file.
    Returns a dict:
      {
        "date_range": "May 1 thru May 3 2026",
        "combined": <bytes of combined xlsx>,
        "companies": {
            "Y&S": <bytes>,
            "Katz": <bytes>,
            ...
        }
      }
    """
    # Accept any sheet name — try Template first, then Sheet1, then first sheet
    xl = pd.ExcelFile(io.BytesIO(file_bytes))
    preferred = ["Template", "Sheet1"]
    sheet = next((s for s in preferred if s in xl.sheet_names), xl.sheet_names[0])
    df_raw = xl.parse(sheet)

    # Determine date range
    dates = fix_date(df_raw["PO Created"]).dt.normalize().dropna().dt.date.unique()
    date_range_str = format_date_range([pd.Timestamp(d) for d in dates])

    # Build query outputs
    all_df = build_all_query(df_raw)
    summary_df = build_summary_query(df_raw)

    # Company filtered sheets
    company_dfs = {
        "Y&S":        filter_company(all_df, ["YS Tickets", "YS-Seatgeek", "YS Tickets Spec", "YS-Seatgeek2"]),
        "Grossman":   filter_company(all_df, ["YSM Tickets"]),
        "Sternbuch":  filter_company(all_df, ["YSS Tickets"]),
        "Pollak":     filter_company(all_df, ["Pollak Tickets"]),
        "Levine":     filter_company(all_df, ["Yoni Levine"]),
        "Levovitz":   filter_company(all_df, ["Levovitz"]),
        "GK":         filter_company(all_df, ["GK LLC"], rename_company={"GK LLC": "YSKG"}),
        "Ticket Guy": filter_company(all_df,
                          ["The Ticket Guy", "The Ticket Guy-Jas", "The Ticket Guy-Legacy", "The Ticket Guy VIP"],
                          vendor_replace={"Broadway Direct": "Box Office - Broadway Inbound"},
                          strip_company_prefix="The "),
        "Chase":      filter_company(all_df, ["Jacks YS"], rename_company={"Jacks YS": "Chase (Jacks)"}),
        "YSA":        filter_company(all_df, ["YSA", "YSA 2", "YSA 3"]),
        "Katz":       filter_company(all_df, ["YS Katz"]),
        "TL":         filter_company(all_df, ["YS TL"]),
        "Waxler":     filter_company(all_df, ["YSW"], rename_company={"YSW": "YSW (Waxler)"}),
        "Damona":       filter_company(all_df, ["Damon and Crew"]),
        "YourTickets":  filter_company(all_df, ["YourTickets"]),
    }

    # ── Build combined workbook ────────────────────────────────────────────────
    wb_combined = openpyxl.Workbook()
    wb_combined.remove(wb_combined.active)
    write_sheet(wb_combined, "Input", df_raw)
    write_sheet(wb_combined, "All", all_df)
    write_sheet(wb_combined, "Summary", summary_df)
    for sheet_name, cdf in company_dfs.items():
        write_sheet(wb_combined, sheet_name, cdf)
        # Red tab for empty sheets
        if len(cdf) == 0:
            wb_combined[sheet_name].sheet_properties.tabColor = "FF0000"

    combined_buf = io.BytesIO()
    wb_combined.save(combined_buf)
    combined_bytes = combined_buf.getvalue()

    # ── Build per-company workbooks (only for non-empty) ─────────────────────────
    company_files = {}
    for sheet_name, cdf in company_dfs.items():
        if len(cdf) == 0:
            continue
        wb = openpyxl.Workbook()
        wb.remove(wb.active)
        write_sheet(wb, sheet_name, cdf)
        buf = io.BytesIO()
        wb.save(buf)
        company_files[sheet_name] = buf.getvalue()

    # All company names (including empty) for UI display
    all_company_names = list(company_dfs.keys())

    # ── Build summary stats for UI display ───────────────────────────────────────
    stats = {}
    # Combined = all_df total
    stats["Combined"] = {
        "rows": len(all_df),
        "total_cost": round(float(all_df["Total Cost"].sum()), 2),
    }
    for sheet_name, cdf in company_dfs.items():
        stats[sheet_name] = {
            "rows": len(cdf),
            "total_cost": round(float(cdf["Total Cost"].sum()), 2) if len(cdf) > 0 else 0.0,
        }

    return {
        "date_range": date_range_str,
        "combined": combined_bytes,
        "companies": company_files,
        "all_companies": all_company_names,
        "stats": stats,
    }
