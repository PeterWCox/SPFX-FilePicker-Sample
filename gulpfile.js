"use strict";
const build = require("@microsoft/sp-build-web");
build.addSuppression(/Warning - /gi);
build.addSuppression(/Warning - lint -/gi); //uncomment for dev work
var getTasks = build.rig.getTasks;
build.rig.getTasks = function ()
{
  var result = getTasks.call(build.rig);
  result.set("serve", result.get("serve-deprecated"));
  return result;
};
/* fast-serve */
const { addFastServe } = require("spfx-fast-serve-helpers");
addFastServe(build);
/* end of fast-serve */
build.initialize(require("gulp"));
build.sass.setConfig({ warnOnNonCSSModules: false, useCssModules: true });
