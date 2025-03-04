
###################################################
dist_list = ['alpha', 'anglit', 'arcsine', 'argus', 'beta', 'betaprime', 'bradford', 'burr', 'burr12',
             'cauchy', 'chi', 'chi2', 'cosine', 'crystalball', 'dgamma', 'dweibull', 'erlang', 'expon',
             'exponnorm', 'exponpow', 'exponweib', 'f', 'fatiguelife', 'fisk', 'foldcauchy', 'foldnorm',
             'frechet_l', 'frechet_r', 'gamma', 'gausshyper', 'genexpon', 'genextreme', 'gengamma',
             'genhalflogistic', 'geninvgauss', 'genlogistic', 'gennorm', 'genpareto', 'gilbrat', 'gompertz',
             'gumbel_l', 'gumbel_r', 'halfcauchy', 'halfgennorm', 'halflogistic', 'halfnorm', 'hypsecant',
             'invgamma', 'invgauss', 'invweibull', 'johnsonsb', 'johnsonsu', 'kappa3', 'kappa4', 'ksone',
             'kstwo', 'kstwobign', 'laplace', 'levy', 'levy_l', 'levy_stable', 'loggamma', 'logistic',
             'loglaplace', 'lognorm', 'loguniform', 'lomax', 'maxwell', 'mielke', 'moyal', 'nakagami',
             'ncf', 'nct', 'ncx2', 'norm', 'norminvgauss', 'pareto', 'pearson3', 'powerlaw', 'powerlognorm',
             'powernorm', 'rayleigh', 'rdist', 'recipinvgauss', 'reciprocal', 'rice', 'semicircular',
             'skewnorm', 't', 'triang', 'truncexpon', 'truncnorm', 'tukeylambda', 'uniform', 'vonmises',
             'vonmises_line', 'wald', 'weibull_max', 'weibull_min', 'wrapcauchy']


#######################################################
dist_parm_dict = {
    "alpha":            ["a", "loc", "scale"],
    "anglit":           ["loc", "scale"],
    "arcsine":          ["loc", "scale"],
    "argus":            ["chi", "loc", "scale"],
    "beta":             ["a", "b", "loc", "scale"],
    "betaprime":        ["a", "b", "loc", "scale"],
    "bradford":         ["c", "loc", "scale"],
    "burr":             ["c", "d", "loc", "scale"],
    "burr12":           ["c", "d", "loc", "scale"],
    "cauchy":           ["loc", "scale"],
    "chi":              ["df", "loc", "scale"],
    "chi2":             ["df", "loc", "scale"],
    "cosine":           ["loc", "scale"],
    "crystalball":      ["beta", "m", "loc", "scale"],
    "dgamma":           ["a", "loc", "scale"],
    "dweibull":         ["c", "loc", "scale"],
    "erlang":           ["a", "loc", "scale"],
    "expon":            ["loc", "scale"],
    "exponnorm":        ["K", "loc", "scale"],
    "exponpow":         ["b", "loc", "scale"],
    "exponweib":        ["a", "c", "loc", "scale"],
    "f":                ["dfn", "dfd", "loc", "scale"],
    "fatiguelife":      ["c", "loc", "scale"],
    "fisk":             ["c", "loc", "scale"],
    "foldcauchy":       ["c", "loc", "scale"],
    "foldnorm":         ["c", "loc", "scale"],
    "frechet_l":        ["c", "loc", "scale"],
    "frechet_r":        ["c", "loc", "scale"],
    "gamma":            ["a", "loc", "scale"],
    "gausshyper":       ["a", "b", "c", "z", "loc", "scale"],
    "genexpon":         ["a", "b", "c", "loc", "scale"],
    "genextreme":       ["c", "loc", "scale"],
    "gengamma":         ["a", "c", "loc", "scale"],
    "genhalflogistic":  ["c", "loc", "scale"],
    "geninvgauss":      ["p", "b", "loc", "scale"],
    "genlogistic":      ["c", "loc", "scale"],
    "gennorm":          ["beta", "loc", "scale"],
    "genpareto":        ["c", "loc", "scale"],
    "gilbrat":          ["loc", "scale"],
    "gompertz":         ["c", "loc", "scale"],
    "gumbel_l":         ["loc", "scale"],
    "gumbel_r":         ["loc", "scale"],
    "halfcauchy":       ["loc", "scale"],
    "halfgennorm":      ["beta", "loc", "scale"],
    "halflogistic":     ["loc", "scale"],
    "halfnorm":         ["loc", "scale"],
    "hypsecant":        ["loc", "scale"],
    "invgamma":         ["a", "loc", "scale"],
    "invgauss":         ["mu", "loc", "scale"],
    "invweibull":       ["c", "loc", "scale"],
    "johnsonsb":        ["a", "b", "loc", "scale"],
    "johnsonsu":        ["a", "b", "loc", "scale"],
    "kappa3":           ["a", "loc", "scale"],
    "kappa4":           ["h", "k", "loc", "scale"],
    "ksone":            ["n", "loc", "scale"],
    "kstwo":            ["n", "loc", "scale"],
    "kstwobign":        ["loc", "scale"],
    "laplace":          ["loc", "scale"],
    "levy":             ["loc", "scale"],
    "levy_l":           ["loc", "scale"],
    "levy_stable":      ["alpha", "beta", "loc", "scale"],
    "loggamma":         ["c", "loc", "scale"],
    "logistic":         ["loc", "scale"],
    "loglaplace":       ["c", "loc", "scale"],
    "lognorm":          ["s", "loc", "scale"],
    "loguniform":       ["a", "b", "loc", "scale"],
    "lomax":            ["c", "loc", "scale"],
    "maxwell":          ["loc", "scale"],
    "mielke":           ["k", "s", "loc", "scale"],
    "moyal":            ["loc", "scale"],
    "nakagami":         ["nu", "loc", "scale"],
    "ncf":              ["dfn", "dfd", "nc", "loc", "scale"],
    "nct":              ["df", "nc", "loc", "scale"],
    "ncx2":             ["df", "nc", "loc", "scale"],
    "norm":             ["loc", "scale"],
    "norminvgauss":     ["a", "b", "loc", "scale"],
    "pareto":           ["b", "loc", "scale"],
    "pearson3":         ["skew", "loc", "scale"],
    "powerlaw":         ["a", "loc", "scale"],
    "powerlognorm":     ["c", "s", "loc", "scale"],
    "powernorm":        ["c", "loc", "scale"],
    "rayleigh":         ["loc", "scale"],
    "rdist":            ["c", "loc", "scale"],
    "recipinvgauss":    ["mu", "loc", "scale"],
    "reciprocal":       ["a", "b", "loc", "scale"],
    "rice":             ["b", "loc", "scale"],
    "semicircular":     ["loc", "scale"],
    "skewnorm":         ["a", "loc", "scale"],
    "t":                ["df", "loc", "scale"],
    "triang":           ["c", "loc", "scale"],
    "truncexpon":       ["b", "loc", "scale"],
    "truncnorm":        ["a", "b", "loc", "scale"],
    "tukeylambda":      ["lam", "loc", "scale"],
    "uniform":          ["loc", "scale"],
    "vonmises":         ["kappa", "loc", "scale"],
    "vonmises_line":    ["kappa", "loc", "scale"],
    "wald":             ["loc", "scale"],
    "weibull_max":      ["c", "loc", "scale"],
    "weibull_min":      ["c", "loc", "scale"],
    "wrapcauchy":       ["c", "loc", "scale"]
}

########################################################
# Removing this three from dictionary due to issues
#    "rv_continuous":    ["None"],
#    "rv_histogram":     ["None"],
#    "trapz":            ["None"],


########################################################
# Old scipy documentation links added in the code if the distributions are [frechet_l, frechet_r and reciprocal]
# trapz is not added in drop down so the link is also ignored

# https://docs.scipy.org/doc/scipy-0.14.0/reference/generated/scipy.stats.frechet_l.html
# https://docs.scipy.org/doc/scipy-0.14.0/reference/generated/scipy.stats.frechet_r.html
# https://docs.scipy.org/doc/scipy-0.14.0/reference/generated/scipy.stats.reciprocal.html
# https://docs.scipy.org/doc/scipy-0.17.0/reference/generated/scipy.integrate.trapz.html


########################################################
# All Distributions

# [‘#alpha’, ‘#anglit’, ‘#arcsine’, ‘#argus’, ‘#beta’, ‘#betaprime’, ‘#bradford’, ‘#burr’, ‘#burr12’, ‘#cauchy’,
# ‘#chi’, ‘#chi2’, ‘#cosine’, ‘#crystalball’, ‘#dgamma’, ‘#dweibull’, ‘#erlang’, ‘#expon’, ‘#exponnorm’, ‘#exponpow’,
# ‘#exponweib’, ‘#f’, ‘#fatiguelife’, ‘#fisk’, ‘#foldcauchy’, ‘#foldnorm’, ‘#frechet_l’, ‘#frechet_r’, ‘#gamma’,
# ‘#gausshyper’, ‘#genexpon’, ‘#genextreme’, ‘#gengamma’, ‘#genhalflogistic’, ‘#geninvgauss’, ‘#genlogistic’,
# ‘#gennorm’, ‘#genpareto’, ‘#gilbrat’, ‘#gompertz’, ‘#gumbel_l’, ‘#gumbel_r’, ‘#halfcauchy’, ‘#halfgennorm’,
# ‘#halflogistic’, ‘#halfnorm’, ‘#hypsecant’, ‘#invgamma’, ‘#invgauss’, ‘#invweibull’, ‘#johnsonsb’,
# ‘#johnsonsu’, ‘#kappa3’, ‘#kappa4’, ‘#ksone’, ‘#kstwo’, ‘#kstwobign’, ‘#laplace’, ‘#levy’, ‘#levy_l’,
# ‘#levy_stable’, ‘#loggamma’, ‘#logistic’, ‘#loglaplace’, ‘#lognorm’, ‘loguniform’, ‘#lomax’, ‘#maxwell’,
# ‘#mielke’, ‘#moyal’, ‘#nakagami’, ‘#ncf’, ‘#nct’, ‘#ncx2’, ‘#norm’, ‘#norminvgauss’, ‘#pareto’,
# ‘#pearson3’, ‘#powerlaw’, ‘#powerlognorm’, ‘#powernorm’, ‘#rayleigh’, ‘#rdist’, ‘#recipinvgauss’,
# ‘#reciprocal’, ‘#rice’, ‘#rv_continuous’, ‘#rv_histogram’, ‘#semicircular’, ‘#skewnorm’, ‘#t’,
# ‘#trapz’, ‘#triang’, ‘#truncexpon’, ‘#truncnorm’, ‘#tukeylambda’, ‘#uniform’, ‘#vonmises’, ‘#vonmises_line’,
# ‘#wald’, ‘#weibull_max’, ‘#weibull_min’, ‘#wrapcauchy’]
