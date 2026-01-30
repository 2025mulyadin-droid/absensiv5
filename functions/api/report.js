export async function onRequestGet(context) {
    const { env, request } = context;
    const url = new URL(request.url);
    const startDate = url.searchParams.get("start");
    const endDate = url.searchParams.get("end");

    if (!startDate || !endDate) {
        return new Response(JSON.stringify({ error: "Start and End date required" }), { status: 400 });
    }

    try {
        // 1. Get detailed list
        const listQuery = `
      SELECT a.date, a.status, u.name 
      FROM attendance a
      JOIN users u ON a.user_id = u.id
      WHERE a.date >= ? AND a.date <= ?
      ORDER BY a.date DESC, u.name ASC
    `;
        const { results: list } = await env.DB.prepare(listQuery)
            .bind(startDate, endDate)
            .all();

        // 2. Get summary per user
        const summaryQuery = `
      SELECT 
        u.name,
        COUNT(CASE WHEN a.status = 'Hadir' THEN 1 END) as total_hadir,
        COUNT(CASE WHEN a.status = 'Sakit' THEN 1 END) as total_sakit,
        COUNT(CASE WHEN a.status = 'Izin' THEN 1 END) as total_izin,
        COUNT(a.id) as total_absen
      FROM users u
      LEFT JOIN attendance a ON u.id = a.user_id AND a.date >= ? AND a.date <= ?
      GROUP BY u.id, u.name
      ORDER BY u.name ASC
    `;
        const { results: summary } = await env.DB.prepare(summaryQuery)
            .bind(startDate, endDate)
            .all();

        return new Response(JSON.stringify({ list, summary }), {
            headers: { "Content-Type": "application/json" },
        });

    } catch (e) {
        return new Response(JSON.stringify({ error: e.message }), { status: 500 });
    }
}
