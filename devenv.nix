{pkgs, ...}: {
  dotenv.enable = false;
  dotenv.disableHint = true;

  languages.javascript.enable = true;
  languages.javascript.corepack.enable = true;
  languages.typescript.enable = true;
}
