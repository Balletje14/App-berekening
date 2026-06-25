// ===============================
// GLOBALE VARIABELEN
// ===============================

let bloemen = [];
let aantallen = {};
let huidigBoeketGewijzigd = false;

// ===============================
// OPSTARTEN
// ===============================

window.onload = function () {

    const opgeslagenBloemen =
        localStorage.getItem("bloemen");

    if (opgeslagenBloemen) {

        bloemen =
            JSON.parse(
                opgeslagenBloemen
            );

        toonBloemen();
    }

    vulBoeketLijst();

};

// ===============================
// EVENTS
// ===============================

document
.getElementById("excelFile")
.addEventListener(
    "change",
    leesExcel
);


// ===============================
// EXCEL INLEZEN
// ===============================

function leesExcel(event) {

    const file =
        event.target.files[0];

    if (!file) return;

    const reader =
        new FileReader();

    reader.onload =
        function (e) {

            const data =
                new Uint8Array(
                    e.target.result
                );

            const workbook =
                XLSX.read(
                    data,
                    {
                        type: "array"
                    }
                );

            const worksheet =
                workbook.Sheets[
                    workbook.SheetNames[0]
                ];

            bloemen =
                XLSX.utils.sheet_to_json(
                    worksheet
                );

            localStorage.setItem(
                "bloemen",
                JSON.stringify(
                    bloemen
                )
            );

            toonBloemen();

            alert(
                bloemen.length +
                " bloemen geladen."
            );
        };

    reader.readAsArrayBuffer(
        file
    );
}

// ===============================
// ZOEKFUNCTIE
// ===============================

function filterBloemen() {

    const zoekElement =
        document.getElementById(
            "zoekveld"
        );

    if (!zoekElement) {
        return;
    }

    const zoekterm =
        zoekElement.value
        .toLowerCase()
        .trim();

    let lijst =
        bloemen;

    if (zoekterm !== "") {

        lijst =
            bloemen.filter(b => {

                const naam =
                    String(
                        b.Naam || ""
                    ).toLowerCase();

                const artikel =
                    String(
                        b.Artikelnummer || ""
                    ).toLowerCase();

                return (
                    naam.includes(
                        zoekterm
                    )
                    ||
                    artikel.includes(
                        zoekterm
                    )
                );

            });

    }

    toonBloemen(
        lijst
    );

}

function leegZoekveld() {

    const zoekveld =
        document.getElementById(
            "zoekveld"
        );

    if (!zoekveld)
        return;

    zoekveld.value = "";

    toonBloemen(
        bloemen
    );

}

// ===============================
// ENTER = TOETSENBORD SLUITEN
// ===============================

function zoekEnter(event) {

    if (
        event.key === "Enter"
    ) {

        filterBloemen();

        event.target.blur();

    }

}

// ===============================
// AANTAL OPSLAAN
// ===============================

function bewaarAantal(
    artikelnummer,
    waarde
) {

    aantallen[
        artikelnummer
    ] =
        parseInt(
            waarde
        ) || 0;

    huidigBoeketGewijzigd =
        true;

    herberekenAutomatisch();

}

// ===============================
// BLOEMEN TONEN
// ===============================

function toonBloemen(
    lijst = bloemen
) {

    let totaalSoorten = 0;
    let totaalStelen = 0;

    let totaalInkoop = 0;
    let totaalVerkoop = 0;
    let totaalBTW = 0;

    bloemen.forEach(b => {

        const aantal =
            aantallen[
                b.Artikelnummer
            ] || 0;

        if (aantal > 0) {

            totaalSoorten++;
            totaalStelen += aantal;

            const inkoopPrijs =
                Number(
                    b["Inkoopprijs"] ??
                    b[" Inkoopprijs "] ??
                    0
                );

            const verkoopPrijs =
                Number(
                    b["Verkoopprijs"] ??
                    b[" Verkoopprijs "] ??
                    0
                );

            const btwPrijs =
                Number(
                    b["BTW"] ??
                    b["Btw"] ??
                    b[" Btw "] ??
                    0
                );

            totaalInkoop +=
                aantal * inkoopPrijs;

            totaalVerkoop +=
                aantal * verkoopPrijs;

            totaalBTW +=
                aantal * btwPrijs;

        }

    });

    let html = `

    <div style="
        margin-bottom:15px;
        padding:15px;
        background:#f5f5f5;
        border-radius:8px;
        border:1px solid #ddd;
    ">

        <div>
            <strong>Geselecteerde soorten:</strong>
            ${totaalSoorten}
        </div>

        <div>
            <strong>Totaal stelen:</strong>
            ${totaalStelen}
        </div>

        <hr>

        <div>
            <strong>Inkoopwaarde:</strong>
            € ${totaalInkoop.toFixed(2)}
        </div>

        <div>
            <strong>Verkoop excl. BTW:</strong>
            € ${totaalVerkoop.toFixed(2)}
        </div>

        <div>
            <strong>BTW:</strong>
            € ${totaalBTW.toFixed(2)}
        </div>

        <div>
            <strong>Verkoop incl. BTW:</strong>
            € ${(totaalVerkoop + totaalBTW).toFixed(2)}
        </div>

    </div>

    <table>

        <tr>

            <th>Artikel</th>

            <th>Naam</th>

            <th>Inkoop p/st</th>

            <th>Verkoop p/st</th>

            <th>BTW p/st</th>

            <th>Aantal</th>

        </tr>

    `;

    lijst.forEach(b => {

        const inkoopPrijs =
            Number(
                b["Inkoopprijs"] ??
                b[" Inkoopprijs "] ??
                0
            );

        const verkoopPrijs =
            Number(
                b["Verkoopprijs"] ??
                b[" Verkoopprijs "] ??
                0
            );

        const btwPrijs =
            Number(
                b["BTW"] ??
                b["Btw"] ??
                b[" Btw "] ??
                0
            );

        html += `

        <tr>

            <td>
                ${b.Artikelnummer || ""}
            </td>

            <td>
                ${b.Naam || ""}
            </td>

            <td>
                € ${inkoopPrijs.toFixed(2)}
            </td>

            <td>
                € ${verkoopPrijs.toFixed(2)}
            </td>

            <td>
                € ${btwPrijs.toFixed(2)}
            </td>

            <td>

    <input
        type="number"
        min="0"

        value="${
            aantallen[
                b.Artikelnummer
            ] || 0
        }"

        oninput="
            bewaarAantal(
                '${b.Artikelnummer}',
                this.value
            );
        "
    >

</td>

        </tr>

        `;

    });

    html += `
    </table>
    `;

    document
        .getElementById(
            "bloemen"
        )
        .innerHTML =
        html;
}
// ===============================
// BEREKEN BOEKET
// ===============================

function bereken() {

    let totaalInkoop = 0;
    let totaalVerkoop = 0;
    let totaalBTW = 0;

    let overzicht = `
    <table>

        <tr>
            <th>Artikel</th>
            <th>Naam</th>
            <th>Aantal</th>
            <th>Inkoop</th>
            <th>Verkoop</th>
            <th>BTW</th>
        </tr>
    `;

    bloemen.forEach(b => {

        const aantal =
            aantallen[
                b.Artikelnummer
            ] || 0;

        if (aantal === 0)
            return;

        const inkoopPrijs =
            Number(
                b["Inkoopprijs"] ??
                b[" Inkoopprijs "] ??
                0
            );

        const verkoopPrijs =
            Number(
                b["Verkoopprijs"] ??
                b[" Verkoopprijs "] ??
                0
            );

        const btwPrijs =
            Number(
                b["BTW"] ??
                b["Btw"] ??
                b[" Btw "] ??
                0
            );

        const regelInkoop =
            aantal *
            inkoopPrijs;

        const regelVerkoop =
            aantal *
            verkoopPrijs;

        const regelBTW =
            aantal *
            btwPrijs;

        totaalInkoop +=
            regelInkoop;

        totaalVerkoop +=
            regelVerkoop;

        totaalBTW +=
            regelBTW;

        overzicht += `

        <tr>

            <td>
                ${b.Artikelnummer}
            </td>

            <td>
                ${b.Naam}
            </td>

            <td>
                ${aantal}
            </td>

            <td>
                € ${regelInkoop.toFixed(2)}
            </td>

            <td>
                € ${regelVerkoop.toFixed(2)}
            </td>

            <td>
                € ${regelBTW.toFixed(2)}
            </td>

        </tr>

        `;

    });

    overzicht += `
    </table>
    `;

    // Overzicht boeket

    document
    .getElementById(
        "boeketOverzicht"
    )
    .innerHTML =
    overzicht;

    // Totalen

    document
    .getElementById(
        "resultaat"
    )
    .innerHTML = `

    <h2>Totalen Boeket</h2>

    <table>

        <tr>
            <td>
                Inkoop excl. BTW
            </td>

            <td>
                € ${totaalInkoop.toFixed(2)}
            </td>
        </tr>

        <tr>
            <td>
                Verkoop excl. BTW
            </td>

            <td>
                € ${totaalVerkoop.toFixed(2)}
            </td>
        </tr>

        <tr>
            <td>
                BTW
            </td>

            <td>
                € ${totaalBTW.toFixed(2)}
            </td>
        </tr>

        <tr>
            <td>
                Verkoop incl. BTW
            </td>

            <td>
                € ${(
                    totaalVerkoop +
                    totaalBTW
                ).toFixed(2)}
            </td>
        </tr>

    </table>

    `;
}

// ===============================
// TOTALEN OPHALEN
// ===============================

function berekenTotalen() {

    let totaalInkoop = 0;
    let totaalVerkoop = 0;
    let totaalBTW = 0;

    bloemen.forEach(b => {

        const aantal =
            aantallen[
                b.Artikelnummer
            ] || 0;

        if (aantal === 0)
            return;

        const inkoopPrijs =
            Number(
                b["Inkoopprijs"] ??
                b[" Inkoopprijs "] ??
                0
            );

        const verkoopPrijs =
            Number(
                b["Verkoopprijs"] ??
                b[" Verkoopprijs "] ??
                0
            );

        const btwPrijs =
            Number(
                b["BTW"] ??
                b["Btw"] ??
                b[" Btw "] ??
                0
            );

        totaalInkoop +=
            aantal *
            inkoopPrijs;

        totaalVerkoop +=
            aantal *
            verkoopPrijs;

        totaalBTW +=
            aantal *
            btwPrijs;

    });

    return {

        totaalInkoop,
        totaalVerkoop,
        totaalBTW,

        totaalInclBTW:
            totaalVerkoop +
            totaalBTW

    };
}

// ===============================
// GESELECTEERDE BLOEMEN
// ===============================

function geselecteerdeBloemen() {

    let lijst = [];

    bloemen.forEach(b => {

        const aantal =
            aantallen[
                b.Artikelnummer
            ] || 0;

        if (aantal > 0) {

            lijst.push({

                Artikelnummer:
                    b.Artikelnummer,

                Naam:
                    b.Naam,

                Aantal:
                    aantal,

                Inkoopprijs:
                    Number(
                        b["Inkoopprijs"] ??
                        b[" Inkoopprijs "] ??
                        0
                    ),

                Verkoopprijs:
                    Number(
                        b["Verkoopprijs"] ??
                        b[" Verkoopprijs "] ??
                        0
                    ),

                BTW:
                    Number(
                        b["BTW"] ??
                        b["Btw"] ??
                        b[" Btw "] ??
                        0
                    )

            });

        }

    });

    return lijst;
}

// ===============================
// NIEUW BOEKET
// ===============================

function nieuwBoeket() {

    if (
        huidigBoeketGewijzigd &&
        !confirm(
            "Niet opgeslagen wijzigingen gaan verloren. Doorgaan?"
        )
    ) {
        return;
    }

    aantallen = {};

    document
        .getElementById("boeketNaam")
        .value = "";

    document
        .getElementById("zoekveld")
        .value = "";

    document
        .getElementById("resultaat")
        .innerHTML = "";

    document
        .getElementById("boeketOverzicht")
        .innerHTML = "";

    huidigBoeketGewijzigd = false;

    toonBloemen();
}

// ===============================
// ALLES LEEGMAKEN
// ===============================

function allesLeegmaken() {

    if (
        confirm(
            "Alle aantallen verwijderen?"
        )
    ) {

        aantallen = {};

        document
            .getElementById(
                "resultaat"
            )
            .innerHTML = "";

        document
            .getElementById(
                "boeketOverzicht"
            )
            .innerHTML = "";

        toonBloemen();

    }

}

// ===============================
// BOEKET OPSLAAN
// ===============================

function opslaanBoeket() {

    const naam =
        document
            .getElementById(
                "boeketNaam"
            )
            .value
            .trim();

    if (!naam) {

        alert(
            "Geef eerst een boeketnaam op."
        );

        return;
    }

    const data = {

        naam: naam,

        datum:
            new Date()
            .toISOString(),

        aantallen:
            aantallen

    };

    localStorage.setItem(
        "boeket_" + naam,
        JSON.stringify(data)
    );

    huidigBoeketGewijzigd = false;

    vulBoeketLijst();

    alert(
        "Boeket opgeslagen."
    );

}

// ===============================
// BOEKET OPENEN
// ===============================

function laadBoeket() {

    const naam =
        document
            .getElementById(
                "opgeslagenBoeketten"
            )
            .value;

    if (!naam)
        return;

    if (
        huidigBoeketGewijzigd
    ) {

        const doorgaan =
            confirm(
                "Er staat mogelijk nog een boeket open.\n\nWeet je zeker dat je een opgeslagen boeket wilt openen?"
            );

        if (!doorgaan)
            return;
    }

    const data =
        JSON.parse(
            localStorage.getItem(
                "boeket_" + naam
            )
        );

    if (!data)
        return;

    aantallen =
        data.aantallen || {};

    document
        .getElementById(
            "boeketNaam"
        )
        .value = naam;

    toonBloemen();

    bereken();

    huidigBoeketGewijzigd =
        false;

}

// ===============================
// BOEKETTEN VULLEN
// ===============================

function vulBoeketLijst() {

    const lijst =
        document
            .getElementById(
                "opgeslagenBoeketten"
            );

    lijst.innerHTML =
        `<option value="">
            Kies opgeslagen boeket
        </option>`;

    Object.keys(localStorage)
        .forEach(key => {

            if (
                key.startsWith(
                    "boeket_"
                )
            ) {

                const naam =
                    key.replace(
                        "boeket_",
                        ""
                    );

                lijst.innerHTML +=
                    `
                    <option value="${naam}">
                        ${naam}
                    </option>
                    `;

            }

        });

}

// ===============================
// BOEKET VERWIJDEREN
// ===============================

function verwijderBoeket() {

    const naam =
        document
            .getElementById(
                "opgeslagenBoeketten"
            )
            .value;

    if (!naam) {

        alert(
            "Selecteer eerst een boeket."
        );

        return;
    }

    const verwijderen =
        confirm(
            `"${naam}" verwijderen?`
        );

    if (!verwijderen)
        return;

    localStorage.removeItem(
        "boeket_" + naam
    );

    vulBoeketLijst();

    alert(
        "Boeket verwijderd."
    );

}

// ===============================
// BOEKET HERNOEMEN
// ===============================

function hernoemBoeket() {

    const oudeNaam =
        document
            .getElementById(
                "opgeslagenBoeketten"
            )
            .value;

    if (!oudeNaam) {

        alert(
            "Selecteer eerst een boeket."
        );

        return;
    }

    const nieuweNaam =
        prompt(
            "Nieuwe naam:",
            oudeNaam
        );

    if (
        !nieuweNaam ||
        nieuweNaam === oudeNaam
    )
        return;

    const data =
        localStorage.getItem(
            "boeket_" +
            oudeNaam
        );

    localStorage.setItem(
        "boeket_" +
        nieuweNaam,
        data
    );

    localStorage.removeItem(
        "boeket_" +
        oudeNaam
    );

    document
        .getElementById(
            "boeketNaam"
        )
        .value =
        nieuweNaam;

    vulBoeketLijst();

    alert(
        "Boeket hernoemd."
    );

}

// ===============================
// RESET APP
// ===============================

function resetApp() {

    const akkoord =
        confirm(
            "ALLE opgeslagen boeketten verwijderen?\n\nDit kan niet ongedaan gemaakt worden."
        );

    if (!akkoord)
        return;

    Object.keys(localStorage)
        .forEach(key => {

            if (
                key.startsWith(
                    "boeket_"
                )
            ) {

                localStorage.removeItem(
                    key
                );

            }

        });

    alert(
        "Alle boeketten verwijderd."
    );

    vulBoeketLijst();

}

// ===============================
// PDF EXPORT
// ===============================

async function maakPDF() {

    const boeketNaam =
        document
            .getElementById(
                "boeketNaam"
            )
            .value
            .trim();

    if (!boeketNaam) {

        alert(
            "Geef eerst een boeketnaam op."
        );

        return;
    }

    const geselecteerd =
        geselecteerdeBloemen();

    if (
        geselecteerd.length === 0
    ) {

        alert(
            "Geen bloemen geselecteerd."
        );

        return;
    }

    const { jsPDF } =
        window.jspdf;

    const doc =
        new jsPDF();

    const vandaag =
        new Date()
            .toLocaleDateString(
                "nl-NL"
            );

    // ===========================
    // Titel
    // ===========================

    doc.setFontSize(18);

    doc.text(
        boeketNaam,
        14,
        20
    );

    doc.setFontSize(11);

    doc.text(
        "Datum: " +
        vandaag,
        14,
        28
    );

    // ===========================
    // Tabel maken
    // ===========================

    const rows = [];

    geselecteerd.forEach(item => {

        rows.push([

            item.Artikelnummer,

            item.Naam,

            item.Aantal,

            "€ " +
            (
                item.Inkoopprijs *
                item.Aantal
            ).toFixed(2),

            "€ " +
            (
                item.Verkoopprijs *
                item.Aantal
            ).toFixed(2),

            "€ " +
            (
                item.BTW *
                item.Aantal
            ).toFixed(2)

        ]);

    });

    doc.autoTable({

        startY: 35,

        head: [[

            "Artikel",

            "Naam",

            "Aantal",

            "Inkoop",

            "Verkoop",

            "BTW"

        ]],

        body: rows,

        theme: "grid",

        styles: {

            fontSize: 9

        },

        headStyles: {

            fillColor:
                [76, 175, 80]

        }

    });

    // ===========================
    // Totalen
    // ===========================

    const totalen =
        berekenTotalen();

    let y =
        doc.lastAutoTable
            .finalY + 15;

    doc.setFontSize(12);

    doc.text(

        "Totaal Inkoop excl. BTW: € " +
        totalen.totaalInkoop
            .toFixed(2),

        14,

        y

    );

    y += 8;

    doc.text(

        "Totaal Verkoop excl. BTW: € " +
        totalen.totaalVerkoop
            .toFixed(2),

        14,

        y

    );

    y += 8;

    doc.text(

        "Totaal BTW: € " +
        totalen.totaalBTW
            .toFixed(2),

        14,

        y

    );

    y += 8;

    doc.text(

        "Totaal Verkoop incl. BTW: € " +
        totalen.totaalInclBTW
            .toFixed(2),

        14,

        y

    );

    // ===========================
    // Bestand opslaan
    // ===========================

    doc.save(
        boeketNaam + ".pdf"
    );

}

// ===============================
// EXCEL EXPORT
// ===============================

function exportExcel() {

    const boeketNaam =
        document
            .getElementById(
                "boeketNaam"
            )
            .value
            .trim();

    if (!boeketNaam) {

        alert(
            "Geef eerst een boeketnaam op."
        );

        return;
    }

    const geselecteerd =
        geselecteerdeBloemen();

    if (
        geselecteerd.length === 0
    ) {

        alert(
            "Geen bloemen geselecteerd."
        );

        return;
    }

    const datum =
        new Date()
            .toLocaleDateString(
                "nl-NL"
            );

    let data = [];

    // ===========================
    // Titelregels
    // ===========================

    data.push([
        "Boeket",
        boeketNaam
    ]);

    data.push([
        "Datum",
        datum
    ]);

    data.push([]);

    // ===========================
    // Kolommen
    // ===========================

    data.push([

        "Artikelnummer",

        "Naam",

        "Aantal",

        "Inkoopprijs",

        "Verkoopprijs",

        "BTW p/st",

        "Totaal Inkoop",

        "Totaal Verkoop",

        "Totaal BTW"

    ]);

    // ===========================
    // Bloemen
    // ===========================

    geselecteerd.forEach(item => {

        const totaalInkoop =
            item.Inkoopprijs *
            item.Aantal;

        const totaalVerkoop =
            item.Verkoopprijs *
            item.Aantal;

        const totaalBTW =
            item.BTW *
            item.Aantal;

        data.push([

            item.Artikelnummer,

            item.Naam,

            item.Aantal,

            item.Inkoopprijs,

            item.Verkoopprijs,

            item.BTW,

            totaalInkoop,

            totaalVerkoop,

            totaalBTW

        ]);

    });

    // ===========================
    // Totalen
    // ===========================

    const totalen =
        berekenTotalen();

    data.push([]);
    data.push([]);

    data.push([
        "TOTAAL",
        "",
        "",
        "",
        "",
        "",
        totalen.totaalInkoop,
        totalen.totaalVerkoop,
        totalen.totaalBTW
    ]);

    data.push([]);

    data.push([
        "Verkoop incl. BTW",
        totalen.totaalInclBTW
    ]);

    // ===========================
    // Worksheet maken
    // ===========================

    const worksheet =
        XLSX.utils.aoa_to_sheet(
            data
        );

    // ===========================
    // Kolombreedtes
    // ===========================

    worksheet["!cols"] = [

        { wch: 18 },
        { wch: 35 },
        { wch: 10 },
        { wch: 15 },
        { wch: 15 },
        { wch: 12 },
        { wch: 15 },
        { wch: 15 },
        { wch: 15 }

    ];

    // ===========================
    // Workbook
    // ===========================

    const workbook =
        XLSX.utils.book_new();

    XLSX.utils.book_append_sheet(

        workbook,

        worksheet,

        "Boeket"

    );

    // ===========================
    // Bestand opslaan
    // ===========================

    XLSX.writeFile(

        workbook,

        boeketNaam + ".xlsx"

    );

}

// ===============================
// LAATST GEOPEND BOEKET
// ===============================

function onthoudBoeket() {

    const naam =
        document
            .getElementById(
                "boeketNaam"
            )
            .value
            .trim();

    if (!naam)
        return;

    localStorage.setItem(
        "laatsteBoeket",
        naam
    );

}

// ===============================
// BOEKET AUTOMATISCH OPENEN
// ===============================

function laadLaatsteBoeket() {

    const naam =
        localStorage.getItem(
            "laatsteBoeket"
        );

    if (!naam)
        return;

    const data =
        localStorage.getItem(
            "boeket_" + naam
        );

    if (!data)
        return;

    const boeket =
        JSON.parse(data);

    aantallen =
        boeket.aantallen || {};

    document
        .getElementById(
            "boeketNaam"
        )
        .value =
        naam;

    toonBloemen();

    bereken();

}

// ===============================
// OPSLAAN + ONTHOUDEN
// ===============================

const origineleOpslaanBoeket =
    opslaanBoeket;

opslaanBoeket =
function () {

    origineleOpslaanBoeket();

    onthoudBoeket();

};

// ===============================
// OPENEN + ONTHOUDEN
// ===============================

const origineleLaadBoeket =
    laadBoeket;

laadBoeket =
function () {

    origineleLaadBoeket();

    onthoudBoeket();

};

// ===============================
// SERVICE WORKER
// ===============================

if (
    "serviceWorker"
    in navigator
) {

    window.addEventListener(
        "load",
        () => {

            navigator
                .serviceWorker
                .register(
                    "./service-worker.js"
                )
                .then(() => {

                    console.log(
                        "Service Worker actief"
                    );

                })
                .catch(err => {

                    console.error(
                        err
                    );

                });

        }
    );

}

// ===============================
// VEILIG OPSTARTEN
// ===============================

window.addEventListener(
    "load",
    () => {

        try {

            vulBoeketLijst();

            if (
                bloemen.length > 0
            ) {

                toonBloemen();

            }

            laadLaatsteBoeket();

        }

        catch (error) {

            console.error(
                error
            );

            alert(
                "Fout tijdens opstarten."
            );

        }

    }
);

// ===============================
// AUTOMATISCH HERBEREKENEN
// ===============================

function herberekenAutomatisch() {

    try {

        bereken();

    }

    catch {

    }

}

// ===============================
// MUTATION OBSERVER
// ===============================

const observer =
    new MutationObserver(
        () => {

            try {

                herberekenAutomatisch();

            }

            catch {

            }

        }
    );

window.addEventListener(
    "load",
    () => {

        const container =
            document.getElementById(
                "bloemen"
            );

        if (
            container
        ) {

            observer.observe(
                container,
                {
                    childList: true,
                    subtree: true
                }
            );

        }

    }
);

// ===============================
// VERSIE
// ===============================

console.log(
    "Boeket Calculator v1.0 geladen"
);
