const { readdirSync, readFileSync, writeFileSync } = require("fs");
let paths = readdirSync("../examples").map(e => `../examples/${e}/README.md`);
let regex = /<!--\s*({(?:\s|(?!-->).)*})\s*-->/;
let metas = paths.map(e => ({ path: e, text: readFileSync(e).toString() }))
    .filter(e => regex.test(e.text))
    .map(e => ({ path: e.path, data: JSON.parse(regex.exec(e.text)[1]) }))
console.log(metas)

let md = `
# \`stdVBA\` Examples

This repository holds examples of using \`stdVBA\`. This should give people a better idea of how to use \`stdVBA\` and libraries.

## Contents

| Title | Tags | Dependencies |
|-------|------|--------------|
${metas.map(function (meta) {
    return "|" + [
        "[" + meta.data.description + "](" + encodeURI(meta.path.substring(3, meta.path.length - "/README.md".length)) + ")",
        meta.data.tags.join(", "),
        meta.data.deps.join(", ")
    ].join("|") + "|"
}).join("\r\n")}
`
writeFileSync("../README.md", md)

