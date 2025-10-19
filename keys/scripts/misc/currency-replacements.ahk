Capslock & x:: {
    currencyNames := "ALL, ANG, AOA, ARS, AUD, AWG, BBD, BDT, BIF, BMD, BND, BOB, BRL, BSD, BTN, BWP, BYR, BZD, CAD, CDF, CHF, CLP, COP, CRC, CUC, CUP, CZK, DJF, DOP, ERN, ETB, EUR, FJD, FKP, GBP, GGP, GHS, GIP, GMD, GNF, GTQ, GYD, HKD, HNL, HTG, HUF, IDR, ILS, INR, JMD, JPY, KES, KHR, KMF, KPW, KRW, KYD, LAK, LKR, LRD, LSL, MDL, MGA, MNT, MOP, MRO, MUR, MWK, MXN, MYR, MZN, NAD, NGN, NIO, NPR, NZD, PEN, PGK, PHP, PKR, PLN, PYG, RON, RWF, SBD, SCR, SGD, SHP, SLL, SOS, SRD, SSP, STD, SZL, THB, TMT, TOP, TTD, TWD, TZS, UAH, UGX, USD, UYU, VEF, VND, VUV, WST, XAF, XCD, XOF, XPF, ZAR, ZMW"
    currencySymbols := "L, ƒ, Kz, $, $, ƒ, $, ৳, Fr, $, $, Bs., R$, $, Nu., P, Br, $, $, Fr, Fr, $, $, ₡, $, $, Kč, Fr, $, Nfk, Br, €, $, £, £, £, ₵, £, D, Fr, Q, $, $, L, G, Ft, Rp, ₪, ₹, $, ¥, Sh, ៛, Fr, ₩, ₩, $, ₭, Rs or රු, $, L, L, Ar, ₮, P, UM, ₨, MK, $, RM, MT, $, ₦, C$, ₨, $, S/., K, ₱, ₨, zł, ₲, lei, Fr, $, ₨, $, £, Le, Sh, $, £, Db, L, ฿, m, T$, $, $, Sh, ₴, Sh, $, $, Bs F, ₫, Vt, T, Fr, $, Fr, Fr, R, ZK"
    currencyNamesArr := StrSplit(currencyNames, ", ")
    currencySymbolsArr := StrSplit(currencySymbols, ", ")
    CreateTextMenuFromStr(currencySymbolsArr, currencyNamesArr)
    Return
}