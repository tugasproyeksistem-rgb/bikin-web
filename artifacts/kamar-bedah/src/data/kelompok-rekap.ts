/* Kelompok utama untuk rekap tab Statistik. Hanya dipakai untuk tampilan
   rekap (nama diagnosa/tindakan yang tersimpan tidak diubah). Aturan dicek
   berurutan dan yang pertama cocok dipakai. Yang tidak cocok masuk "Lainnya".
   Kalau ada nama masuk kelompok yang salah, ubah atau tambah kata kunci di sini. */

export const KELOMPOK_TINDAKAN: Array<[string, RegExp]> = [
  ["Katarak (SICS / IOL / Fako)", /\bsics\b|\biol\b|\bfako|\bphaco|\bkatarak|\bekstraksi\b.*\blensa|\bice\b/],
  ["Obstetri (Sectio Caesarea)", /\bsc\b|\bsectio|\bcaesar|\bcesar/],
  ["Adenotonsilektomi & THT", /\bate\b|\bt[oa]nsil|\badeno|\bturbinoplast|\bseptoplast|\bkonka|\bcwl\b|\bantrostom|\bfess\b|\bcaldwell|\bpolip nasi|\bpolypectomy nasi|\brhino|\bsinus|\bantrotom/],
  ["Debridement", /\bdebride|\bdebridem|\bnecrotom|\bnekrotom/],
  ["Appendektomi", /\bap+e?n?d[ie]c?[ct]?[ct]?omi|\bappendec|\bapendec|\bappendict|\bapendict|\bappendix/],
  ["Hemoroidektomi", /\bh?a?e?mo?r+h?o?i?d+e?c?t?omi|\bhaemor|\bhemor|\bhaemo/],
  ["Hernia (Herniorafi / Hernia repair)", /\bhern|\bhermia|\bhild\b/],
  ["Removal implan (plate, screw, K-wire)", /\bremoval|\baff\b|\bangkat (plate|screw|implan|k)|\blepas implan|\bimplant removal/],
  ["ORIF / Fiksasi tulang (plating, pinning, screw)", /\borif\b|\bplat+ing|\bpin+ing|\bpining|\bk[- ]?wire|\bscrew|\btbw\b|\btension band|\btransfixing|\bfiksasi|\bfixation|\bcrif\b|\bopen reduc|\bnailing|\bimplan/],
  ["Reposisi / Gips", /\breposisi|\bgyps|\bgips|\breduksi tertutup|\bclosed reduc|\bbackslab|\bslab/],
  ["Amputasi", /\bamputasi|\bdisartikulasi|\bexartikulasi/],
  ["Laparotomi, Kolesistektomi & Bedah digestif", /\blap[ao]r?[ao]?t|\bchol?[ei]?[sc]?[iy]?stectom|\bcholecyst|\bkolesist|\bcholesist|\bkolostomi|\bcolostomy|\bhemikolektomi|\bhemicolect|\bdigestif|\badhesiolisis|\bgastrectomy|\bileostomi|\bfistulectom|\bfistullectom|\bfistulotomi|\blaparoscop/],
  ["Ginekologi (TAH, BSO, Kuretase, Kistektomi)", /\btah\b|\bbso\b|\buso\b|\bcurret|\bkuretase|\bcuret|\bkistektomi|\bkistectomy|\bcystectomy|\bmiomektomi|\bmyomectomy|\bleep\b|\bhisterektomi|\bsalpingo|\bkonisasi|\bpolip endometrium|\biud\b/],
  ["Insisi & Drainase", /\binsisi|\bincisi|\bincision|\bdrainage|\bdarinage|\bdrainase|\babs[ce]?s/],
  ["Eksisi, Ekstirpasi & Biopsi (tumor, kista, benjolan)", /\bex?s?[ck]?[ie]?[sz]?[sx]?[ie]?[cs]?i|\beksisi|\bexsisi|\bexici|\bexcisi|\beksterp|\bexterp|\bekterp|\bekstirp|\bextirp|\bbiopsi|\bbiopsy|\bfam\b|\blumpectomy|\bmastectomy|\bmastektomi|\bkistektomi|\btumor|\blipoma|\batheroma|\bganglion|\bnodul|\bwide exc|\bwle\b|\bpa\b/],
  ["Sirkumsisi", /\bsirkumsisi|\bsircumsisi|\bcircumcis|\bsirkum|\bsircum/],
  ["Tendon / Bedah tangan (release, repair)", /\brelease|\bthenor|\btenor|\btendo|\btenorafi|\btenorrhaph|\bcarpal tunnel|\bctr\b|\btrigger finger/],
  ["Tiroid (Lobektomi / Tiroidektomi)", /\blobectomy|\bisthmo|\bismo|\bismolob|\bthyroid|\btiroid|\bstruma/],
  ["Urologi", /\bturp\b|\bsistostomi|\bcystostomy|\bnefrektomi|\bureter|\bbuli|\bvasektomi|\bhidrokel|\bhydrocele|\borchi|\bsistoskopi|\bcystoscop|\bdj stent/],
];

export const KELOMPOK_DIAGNOSA: Array<[string, RegExp]> = [
  ["Katarak", /\bkatarak|\bcatarac|\bcataract|\bkatarax/],
  ["Appendicitis", /\bapp?\b|\bappendi|\bapendi|\bap+e?n?d[ie]c?i?[ts]i?s|\bappendic|\bapendic|\bappendec|\bappendisitis|\bappendisitis/],
  ["Tumor / Fibroadenoma Mamma", /\bmamm?ae?\b|\bmamma\b|\bfam\b|\bca mamm|\bpayudara|\bfibroadenoma|\bfibroadenom/],
  ["Kolelitiasis / Kolesistitis", /\bch?o?l[ei]?[sc]?[yi]?s?t?[eo]?li?th?i?a?s[ie]?s|\bcholelit|\bcholelith|\bkolelit|\bcholecystitis|\bkolesistitis|\bbatu empedu|\bgb stone|\bgallstone|\bclolel/],
  ["Hipertrofi Adenotonsil / Tonsilitis", /\badenotonsil|\btonsil|\badenoid|\bhipertrofi adeno|\bhypertrofi adeno|\bate\b|\btosil/],
  ["THT lain (sinusitis, polip, septum, konka)", /\bsinusitis|\bpolip nasi|\bpolyp|\bturbinat|\bethmoid|\bnasal|\bpansinus|\brhinosinus|\bdeviasi septum|\bseptum|\bkonka|\brhinitis|\brinitis|\bepistaksis|\bfraktur os nasal|\bos nasal|\bcwl|\bkolesteatoma/],
  ["Fraktur / Union (kelainan tulang)", /(^|\s)#|\bof digit|\bfraktur|\bfracture|\bfx\b|\bcf\b|\bef\b|\bclosed fr|\bopen fr|\bunion|\bnon union|\bmalunion|\bdislokasi|\bdislocation|\bmaleolus|\bclavicula|\bklavikula|\bfemur|\btibia|\bradius|\bulna|\bhumerus|\bmetacarpal|\bmetatarsal|\bmandibula|\bmaxila|\bpatella|\bcalcaneus|\bacl\b|\bmeniscus|\bhardware|\bimplant/],
  ["Ulkus / Kaki diabetik / Gangren", /\bulk[uo]s|\bulcus|\bdm pedis|\bdiabetic foot|\bgangg?ren|\bgangrene|\bnekrosis|\bnecrosis|\bburn|\bluka bakar|\bluka\b|\bvulnus|\bdekubitus|\bgangraen/],
  ["Abses / Infeksi jaringan lunak", /\babs[ce]?[se]?s|\babcess|\bselulitis|\bcellulitis|\bfurunkel|\bfurunkul|\bcarbunkel|\binfeksi|\binfection|\bbisul|\bparonikia|\bfistul|\bfistel|\bemph?[iy]ema|\bpilonidal/],
  ["Hemoroid", /\bh?a?e?mo?r+h?o?i?d|\bhaemor|\bhemor|\bhm grade|\bwasir/],
  ["Hernia", /\bhern|\bhermia|\bhil[sd]?\b|\bhiel\b|\bhidl\b|\bhsd\b|\bhumb/],
  ["Peritonitis / Akut abdomen", /\bperitonitis|\bileus|\bobstruksi|\bperforasi|\bperforation|\bsubileus|\bakut abdomen|\bacute abdomen|\bvolvulus|\binvaginasi|\bintussus/],
  ["Ginekologi / Obstetri", /\bmyoma|\bmioma|\bkista ovarium|\bovarium|\bbsc\b|\bcin\b|\bgemeli|\bmalpresentasi|\busia tua|\bpost date|\bletak obl|\briwayat sc|\bpartus|\bhamil|\bgravid|\bg ?\d|\bkehamilan|\bca cervi|\bca serviks|\bendometri|\bpolip endo|\buterus|\bhpp\b|\bket\b|\bsolusio|\bplasenta|\bpre ?eklamp|\bpeb\b|\bkpd\b|\bletak sungsang|\bcpd\b|\bserotinus|\bpostdate|\bmenorrag|\bhiper?menorea|\bdub\b|\babortus|\bmola|\bfaetal|\bblighted|\bbligted|\buso\b/],
  ["Tumor / Benjolan jaringan lunak & Limfadenopati", /\btumor|\blipoma|\batheroma|\bkista|\bcyst|\bganglion|\bnodul|\blimphadenop|\blymphadenop|\blympadenop|\blimfadenop|\bbenjolan|\bmassa|\bkeloid|\bfibroma|\bneoplasma|\bca\b|\bcarcinoma|\bkarsinoma|\bmalignan|\bpolip|\bsoft tissue|\bsts\b|\blimpadenop|\bnevus/],
  ["Tiroid / Struma", /\bstruma|\bthyroid|\btiroid|\bnodul tiroid|\bgoiter|\bgoitre|\bstrum/],
  ["Urologi", /\bbph\b|\bbatu ginjal|\bbatu buli|\bbatu ureter|\bnefrolit|\bstriktur|\bphimosis|\bfimosis|\bparafimosis|\bparfimosis|\bhidrokel|\bhydrocele|\bretensi urin|\bhematuri|\bhematuria|\burolit|\bsistitis|\bureter|\bvesicolith|\btorsio testis|\bkriptorkid|\bundescended|\btorsi testis|\btorsio/],
  ["Tangan (CTS, trigger finger)", /\bcts\b|\bcarpal tunnel|\btrigger finger|\bdupuytren|\bde quervain|\btrig+er|\bthumb/],
  ["Mata lain (pterigium)", /\bpter[yi]gium|\bptereqium/],
  ["Trauma / Cedera lain", /\bcorpal|\bcorpus alien|\btrauma|\bvulnus|\bcedera|\binjury|\blaserasi|\bvulnus|\bluka|\bdislokasi|\bcontusio|\bkontusio|\bamputasi|\btendon|\btendo|\bruptur|\brupture|\bsprain|\bligament|\bluxatio|\bdislocatio/],
];

const cari = (nama: string, aturan: Array<[string, RegExp]>): string => {
  const s = String(nama || "").replace(/\s+/g, " ").trim().toLowerCase();
  for (const [n, re] of aturan) { if (re.test(s)) return n; }
  return "Lainnya";
};

export const kelompokTindakan = (nama: string): string => cari(nama, KELOMPOK_TINDAKAN);
export const kelompokDiagnosa = (nama: string): string => cari(nama, KELOMPOK_DIAGNOSA);
