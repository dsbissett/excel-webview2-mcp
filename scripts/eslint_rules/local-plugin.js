import enforceZodSchemaRule from './enforce-zod-schema-rule.js';
import noDirectThirdPartyImportsRule from './no-direct-third-party-imports-rule.js';

export default {
  rules: {
    'no-direct-third-party-imports': noDirectThirdPartyImportsRule,
    'enforce-zod-schema': enforceZodSchemaRule,
  },
};
