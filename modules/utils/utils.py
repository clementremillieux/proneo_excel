"""_summary_

Returns:
    _type_: _description_
"""

from datetime import datetime


def get_current_date_hour() -> str:
    """_summary_

    Returns:
        str: _description_
    """
    jours_semaine = {
        'monday': 'lundi',
        'tuesday': 'mardi',
        'wednesday': 'mercredi',
        'thursday': 'jeudi',
        'friday': 'vendredi',
        'saturday': 'samedi',
        'sunday': 'dimanche'
    }

    mois_annee = {
        'january': 'janvier',
        'february': 'février',
        'march': 'mars',
        'april': 'avril',
        'may': 'mai',
        'june': 'juin',
        'july': 'juillet',
        'august': 'août',
        'september': 'septembre',
        'october': 'octobre',
        'november': 'novembre',
        'december': 'décembre'
    }

    date_actuelle = datetime.now()

    jour_semaine = date_actuelle.strftime('%A')

    mois = date_actuelle.strftime('%B')

    jour_semaine_fr = jours_semaine.get(jour_semaine.lower(), jour_semaine)

    mois_fr = mois_annee.get(mois.lower(), mois)

    date_formatee = date_actuelle.strftime(
        f'{jour_semaine_fr} %d {mois_fr} à %Hh%Mm%Ss')

    return date_formatee
