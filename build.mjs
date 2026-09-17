import fs from 'node:fs';
const config=JSON.parse(fs.readFileSync('data.json','utf8'));
const candidates=config.candidatesData;
const html=fs.readFileSync('template.html','utf8').replace('<!-- APP_DATA -->','<script id="app-data" type="application/json">'+JSON.stringify(config,null,2).replaceAll('<','\\u003c')+'</script>');
fs.writeFileSync('Index.html',html);
console.log(Object.fromEntries(Object.entries(candidates).map(([k,v])=>[k,v.length])));
