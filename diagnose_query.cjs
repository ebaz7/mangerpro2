const fetch = require('node-fetch');

const SAYAN_URL = "http://lep.templatetesti.shop:5000/api/external/v1/query";
const SAYAN_KEY = "s_gate_live_vzje5nkn7q4u";

const sql = `
    WITH GroupedStock AS (
        SELECT 
            t11.Field_005 as ItemCode,
            SUM(CASE 
                WHEN RTRIM(LTRIM(t10.Field_009)) IN ('10', '24', '26', '29', '40', '44', '46', '83') THEN t11.Field_006 
                WHEN RTRIM(LTRIM(t10.Field_009)) IN ('23', '25', '30', '37', '42', '84', '62', '68', '71', '74', '80') THEN -t11.Field_006 
                ELSE 0 
            END) as StockQty
        FROM STR_TBL_011 t11 WITH (NOLOCK)
        INNER JOIN STR_TBL_010 t10 WITH (NOLOCK) ON t11.Field_004 = t10.Field_005 
                                                 AND t11.Field_003 = t10.Field_004 
                                                 AND t11.Field_012 = t10.Field_018
        WHERE t10.Field_008 <= '2025-03-20T23:59:59.000Z'
        GROUP BY t11.Field_005
    )
    SELECT TOP 50
        gs.ItemCode,
        gs.StockQty,
        COALESCE(
            NULLIF(RTRIM(LTRIM(s04.Field_003)), ''),
            NULLIF(RTRIM(LTRIM(t22.Field_004)), ''),
            NULLIF(RTRIM(LTRIM(t02_exact.Field_003)), ''),
            NULLIF(RTRIM(LTRIM(t_name.ItemName)), ''),
            NULLIF(RTRIM(LTRIM(t_group.GroupName)), ''),
            NULLIF(RTRIM(LTRIM(c01.Field_003)), ''),
            RTRIM(LTRIM(gs.ItemCode)),
            N'کالای بدون نام'
        ) as ItemName,
        t_group.GroupName,
        t_group.SubGroupName
    FROM GroupedStock gs
    LEFT JOIN STR_TBL_004 s04 WITH (NOLOCK) ON s04.Field_004 = gs.ItemCode
    LEFT JOIN IND_TBL_022 t22 WITH (NOLOCK) ON t22.Field_005 = gs.ItemCode
    LEFT JOIN IND_TBL_002 t02_exact WITH (NOLOCK) ON t02_exact.Field_008 = gs.ItemCode
    LEFT JOIN COM_TBL_001 c01 WITH (NOLOCK) ON c01.Field_004 = gs.ItemCode
    LEFT JOIN (
        SELECT t21_sub.Field_004 as ItemCode, MIN(t02_sub.Field_003) as ItemName
        FROM IND_TBL_021 t21_sub WITH (NOLOCK)
        LEFT JOIN IND_TBL_002 t02_sub WITH (NOLOCK) ON t21_sub.Field_003 = t02_sub.Field_008
        GROUP BY t21_sub.Field_004
    ) t_name ON gs.ItemCode = t_name.ItemCode
    LEFT JOIN (
        SELECT t21_sub.Field_004 as ItemCode, 
               MIN(t02_sub.Field_003) as SubGroupName,
               MIN(COALESCE(t02_grandparent.Field_003, t02_parent.Field_003, t02_sub.Field_003)) as GroupName
        FROM IND_TBL_021 t21_sub WITH (NOLOCK)
        LEFT JOIN IND_TBL_002 t02_sub WITH (NOLOCK) ON t21_sub.Field_003 = t02_sub.Field_008
        LEFT JOIN IND_TBL_002 t02_parent WITH (NOLOCK) ON t02_sub.Field_009 = t02_parent.Field_008
        LEFT JOIN IND_TBL_002 t02_grandparent WITH (NOLOCK) ON t02_parent.Field_009 = t02_grandparent.Field_008
        GROUP BY t21_sub.Field_004
    ) t_group ON gs.ItemCode = t_group.ItemCode
`;

async function main() {
    console.log("Running Sayan Warehouse Inventory query with TOP 50 and NOLOCK...");
    const startTime = Date.now();
    try {
        const res = await fetch(SAYAN_URL, {
            method: 'POST',
            headers: {
                'Authorization': `Bearer ${SAYAN_KEY}`,
                'Content-Type': 'application/json'
            },
            body: JSON.stringify({ query: sql }),
            timeout: 25000
        });
        
        const duration = Date.now() - startTime;
        console.log(`HTTP Status: ${res.status} (took ${duration}ms)`);
        const data = await res.json();
        if (data.success) {
            console.log(`Success! Retrieved ${data.data.length} items.`);
            console.log("Sample items:", JSON.stringify(data.data.slice(0, 2), null, 2));
        } else {
            console.error("Sayan SQL Error:", data.error || data.message);
        }
    } catch (err) {
        console.error("Fetch failed:", err.message);
    }
}

main();
