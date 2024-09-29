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
        'Monday': 'lundi',
        'Tuesday': 'mardi',
        'Wednesday': 'mercredi',
        'Thursday': 'jeudi',
        'Friday': 'vendredi',
        'Saturday': 'samedi',
        'Sunday': 'dimanche'
    }

    mois_annee = {
        'January': 'janvier',
        'February': 'février',
        'March': 'mars',
        'April': 'avril',
        'May': 'mai',
        'June': 'juin',
        'July': 'juillet',
        'August': 'août',
        'September': 'septembre',
        'October': 'octobre',
        'November': 'novembre',
        'December': 'décembre'
    }

    date_actuelle = datetime.now()

    jour_semaine = date_actuelle.strftime('%A')

    mois = date_actuelle.strftime('%B')

    jour_semaine_fr = jours_semaine.get(jour_semaine.lower(), jour_semaine)

    mois_fr = mois_annee.get(mois.lower(), mois)

    date_formatee = date_actuelle.strftime(
        f'{jour_semaine_fr} %d {mois_fr} à %Hh%Mm%Ss')

    return date_formatee
