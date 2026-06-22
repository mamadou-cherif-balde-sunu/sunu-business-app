# relation_client/utils_report.py
import calendar
from datetime import date, timedelta


def get_week_range(d: date):
    """Retourne (debut, fin) d'une semaine glissante de 7 jours pile,
    ancrée sur le lundi de la semaine ISO contenant d."""
    monday = d - timedelta(days=d.weekday())  # weekday(): lundi=0
    sunday = monday + timedelta(days=6)
    return monday, sunday


def get_month_range(year: int, month: int):
    """Retourne (debut, fin) du mois, en tenant compte automatiquement
    du nombre de jours réel (28/29/30/31) via calendar.monthrange."""
    last_day = calendar.monthrange(year, month)[1]  # gère février bissextile
    return date(year, month, 1), date(year, month, last_day)


def label_week(debut: date, fin: date) -> str:
    mois_fr = ["", "Jan.", "Fév.", "Mars", "Avr.", "Mai", "Juin",
               "Juil.", "Août", "Sept.", "Oct.", "Nov.", "Déc."]
    if debut.month == fin.month:
        return f"Semaine du {debut.day:02d} au {fin.day:02d} {mois_fr[fin.month]} {fin.year}"
    return f"Semaine du {debut.day:02d} {mois_fr[debut.month]} au {fin.day:02d} {mois_fr[fin.month]} {fin.year}"


def label_month(year: int, month: int) -> str:
    mois_fr = ["", "Janvier", "Février", "Mars", "Avril", "Mai", "Juin",
               "Juillet", "Août", "Septembre", "Octobre", "Novembre", "Décembre"]
    return f"{mois_fr[month]} {year}"
# --- à ajouter dans relation_client/utils_report.py ---

def compute_stats(queryset):
    """Calcule NPS, CSAT, CES sur un queryset de Reponse déjà filtré.
    Retourne un dict prêt à l'emploi pour le PPTX et le Dashboard."""
    total = queryset.count()
    if total == 0:
        return {
            "total": 0, "promoteurs": 0, "passifs": 0, "detracteurs": 0,
            "nps": 0, "satisfaits": 0, "insatisfaits": 0, "csat": 0,
            "effort_faible": 0, "effort_modere": 0, "effort_eleve": 0, "ces": 0,
        }

    promoteurs = queryset.filter(categorie_nps="Promoteur").count()
    passifs = queryset.filter(categorie_nps="Passif").count()
    detracteurs = queryset.filter(categorie_nps="Detracteur").count()

    satisfaits = queryset.filter(csat="Satisfait").count()
    insatisfaits = queryset.filter(csat="Insatisfait").count()

    effort_faible = queryset.filter(ces="Effort faible").count()
    effort_modere = queryset.filter(ces="Effort modere").count()
    effort_eleve = queryset.filter(ces="Effort eleve").count()

    nps = round(((promoteurs - detracteurs) / total) * 100, 1)
    csat = round((satisfaits / total) * 100, 1) if (satisfaits + insatisfaits) > 0 else 0
    ces = round(((effort_eleve - effort_faible) / total) * 100, 1)

    return {
        "total": total,
        "promoteurs": promoteurs, "passifs": passifs, "detracteurs": detracteurs,
        "nps": nps,
        "satisfaits": satisfaits, "insatisfaits": insatisfaits, "csat": csat,
        "effort_faible": effort_faible, "effort_modere": effort_modere,
        "effort_eleve": effort_eleve, "ces": ces,
    }


def get_period_stats(Reponse, debut, fin):
    """Filtre les réponses sur [debut, fin] inclus par date_appel,
    et retourne les stats VIE / IARD / GLOBAL."""
    qs = Reponse.objects.filter(date_appel__gte=debut, date_appel__lte=fin)
    qs_vie = qs.filter(canal__icontains="VIE")
    qs_iard = qs.filter(canal__icontains="IARD")

    return {
        "vie": compute_stats(qs_vie),
        "iard": compute_stats(qs_iard),
        "global": compute_stats(qs),
    }
    # --- à ajouter dans relation_client/utils_report.py ---

def interpretation_nps(nps: float) -> str:
    if nps >= 50:
        return (f"avec un NPS de {nps}%, les clients sondés affichent une forte propension "
                f"à recommander SUNU Assurances à leurs proches, amis et connaissances")
    elif nps >= 0:
        return (f"avec un NPS de {nps}%, les clients sondés restent globalement favorables "
                f"à SUNU Assurances, mais la proportion de détracteurs reste à surveiller")
    else:
        return (f"avec un NPS de {nps}%, le nombre de détracteurs dépasse celui des promoteurs, "
                f"ce qui appelle une action correctrice rapide sur la qualité de service")


def interpretation_csat(csat: float) -> str:
    if csat >= 90:
        return f"{csat}% des clients interrogés se déclarent satisfaits du traitement de leurs requêtes"
    elif csat >= 70:
        return (f"{csat}% des clients interrogés se déclarent satisfaits du traitement de leurs requêtes, "
                f"un niveau correct mais perfectible")
    else:
        return (f"seuls {csat}% des clients interrogés se déclarent satisfaits, "
                f"un niveau de satisfaction préoccupant qui nécessite une attention particulière")


def interpretation_ces(ces: float) -> str:
    if ces <= -50:
        return (f"avec un CES de {ces}%, la majorité des clients estiment avoir fourni un effort "
                f"faible pour accéder à l'information recherchée")
    elif ces <= 0:
        return (f"avec un CES de {ces}%, l'effort perçu par les clients pour accéder à "
                f"l'information reste globalement maîtrisé")
    else:
        return (f"avec un CES de {ces}%, une part importante des clients estime avoir fourni "
                f"un effort élevé pour accéder à l'information, ce qui doit être corrigé")


def conclusion_perimetre(nom_perimetre: str, stats: dict) -> str:
    """Génère la conclusion finale (slides VIE/IARD) selon les seuils."""
    nps, csat = stats["nps"], stats["csat"]
    if nps >= 50 and csat >= 80:
        return (f"Les indicateurs CES et CSAT confirment la confiance que nos clients placent en "
                f"SUNU Assurances sur le périmètre {nom_perimetre}. Cette satisfaction durable renforce "
                f"notre engagement et transforme nos assurés en partenaires fidèles, prêts à recommander "
                f"notre compagnie à leurs proches.")
    elif nps >= 0 and csat >= 50:
        return (f"Les indicateurs CES et CSAT témoignent d'un niveau de satisfaction correct sur le "
                f"périmètre {nom_perimetre}. Cependant, un point d'attention particulier devrait être "
                f"apporté aux clients détracteurs et insatisfaits afin de consolider cette dynamique.")
    else:
        return (f"Les indicateurs CES et CSAT révèlent une fragilité de la satisfaction client sur le "
                f"périmètre {nom_perimetre}. Une action correctrice prioritaire est recommandée afin de "
                f"limiter le risque de détérioration de la relation client.")