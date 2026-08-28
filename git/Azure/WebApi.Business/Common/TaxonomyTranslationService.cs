using Microsoft.Extensions.Caching.Memory;
using Microsoft.Extensions.Logging;
using Microsoft.Graph.Models.TermStore;
using Contoso.Portal.Data.DAL.Helpers;

namespace Contoso.Portal.Common
{
    public class TaxonomyTranslationService : ServiceBasePnP<TaxonomyTranslationService>
    {
        private readonly IMemoryCache _memoryCache;

        private readonly string _taxonomySiteId;
        private readonly string _appTermGroup;
        private readonly string _defaultLocale;
        private readonly int _cacheExpirationHours;

        public TaxonomyTranslationService(IMemoryCache memoryCache, ILogger<TaxonomyTranslationService> logger,
            M365AuthHelper auth, string defaultLocale, string taxonomySiteId, string appTermGroup, int cacheExpirationHours = 4)
            : base(logger, auth)
        {
            _memoryCache = memoryCache;
            _defaultLocale = defaultLocale;
            _cacheExpirationHours = cacheExpirationHours;
            _taxonomySiteId = taxonomySiteId;
            _appTermGroup = appTermGroup;
        }

        public async Task<Dictionary<Guid, string>> GetLabelsByIdAsync(Guid[] termsIds, string? locale)
        {
            var taxonomy = await GetTaxonomyCache();
            var result = new Dictionary<Guid, string>();

            foreach (var termId in termsIds)
            {
                taxonomy.TryGetValue(termId, out var term);
                if (term == null)
                    continue; // TODO: add log warning

                var labelLocale = string.IsNullOrECNTy(locale) ? _defaultLocale : locale;
                var label = term?.Labels?.FirstOrDefault((l) => l.LanguageTag == labelLocale);

                // if specific label does not exist use term default one.
                if (label == null)
                    label = term?.Labels?.FirstOrDefault((l) => l.IsDefault == true); // todo: add logger warning

                if (string.IsNullOrECNTy(label?.Name))
                    continue; // todo: add logger warning

                result.TryAdd(termId, label!.Name!);
            }

            return result;
        }

        private async Task<Dictionary<Guid, Term>> GetTaxonomyCache()
        {
            var cacheKey = _taxonomySiteId + _appTermGroup + _defaultLocale;

            return await _memoryCache.GetOrCreateAsync(cacheKey, async (entry) =>
            {
                entry.AbsoluteExpirationRelativeToNow = TimeSpan.FromHours(_cacheExpirationHours);

                using var systemContext = await CreatePnPContextAsSystem();
                using var graphHelper = new GraphHelper(systemContext);
                return await graphHelper.GetAllTermsFrom(_taxonomySiteId, _appTermGroup);
            });
        }
    }

} //ns