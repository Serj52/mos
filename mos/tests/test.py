import unittest
from .mos.moex import rec_excl, curr_pars, curr_pars_ver2



class TestMos(unittest.TestCase):
    path = 'https://moex.com/export/derivatives/currency-rate.aspx?'
    currency_usd = 'USD_RUB'
    params_usd = {'language': 'en', 'currency': currency_usd, 'moment_start': '2021-07-09',
              'moment_end': '2021-08-09'}
    currency_eur = 'EUR_RUB'
    params_eur = {'language': 'en', 'currency': currency_eur, 'moment_start': '2021-07-09',
                  'moment_end': '2021-08-09'}

    def test_moex(self):
        # Проверка количества дней и кол-ва записей в файле Excel
        self.assertEqual(rec_excl(dict_curr=curr_pars(self.path, self.params_usd), currency=self.currency_usd),
                         len(curr_pars(self.path, self.params_usd)))
        self.assertEqual(rec_excl(dict_curr=curr_pars(self.path, self.params_usd), currency=self.currency_usd),
                         len(curr_pars_ver2(self.path, self.params_usd)))

        self.assertEqual(rec_excl(dict_curr=curr_pars(self.path, self.params_eur), currency=self.currency_eur),
                         len(curr_pars(self.path, self.params_eur)))
        self.assertEqual(rec_excl(dict_curr=curr_pars(self.path, self.params_eur), currency=self.currency_eur),
                         len(curr_pars_ver2(self.path, self.params_eur)))

        # Проверка равенства записей курсов валют в файле Excel
        self.assertEqual(rec_excl(dict_curr=curr_pars(self.path, self.params_usd), currency=self.currency_usd),
                         rec_excl(dict_curr=curr_pars(self.path, self.params_eur), currency=self.currency_eur))




