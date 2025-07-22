from ac_price_module import ac_price
from mouser_module import mouser_price
from octopart_module import octopart_price_nexar

def get_online_results(artikelnummer):
    results = []
    if artikelnummer and isinstance(artikelnummer, str) and artikelnummer and artikelnummer.lower() != 'nan':
        ac_res = ac_price(artikelnummer)
        if ac_res:
            results.append(ac_res)
        mouser_res = mouser_price(artikelnummer)
        if mouser_res:
            results.append(mouser_res)
        octopart_res = octopart_price_nexar(artikelnummer)
        if octopart_res:
            results.append(octopart_res)
    return results
