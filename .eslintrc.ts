import { off } from "gulp";
import { error } from "jquery";

require('@rushstack/eslint-config/patch/modern-module-resolution');
module.exports = {
    extends: ['@microsoft/eslint-config-spfx/lib/profiles/react'],
    parserOptions: { tsconfigRootDir: __dirname },
    overrides: [
        {
            files: ["*.test.tsx", "*.tsx"],
            "rules": {
                "class-name": off,
                "export-name": off,
                "forin": off,
                "label-position": off,
                "member-access": error,
                "no-arg": off,
                "no-console": off,
                "no-construct": off,
                "no-duplicate-variable": error,
                "no-eval": off,
                "no-function-expression": error,
                "no-internal-module": error,
                "no-shadowed-variable": error,
                "no-switch-case-fall-through": error,
                "no-unnecessary-semicolons": error,
                "no-unused-expression": error,
                "no-with-statement": error,
                "semicolon": error,
                "trailing-comma": off,
                "typedef": off,
                "typedef-whitespace": off,
                "use-named-parameter": error,
                "variable-name": off,
                "whitespace": off
            }
        }
    ]
};
