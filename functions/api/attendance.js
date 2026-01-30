export async function onRequestGet(context) {
    const { env, request } = context;
    const url = new URL(request.url);
    const dateStr = url.searchParams.get("date"); // YYYY-MM-DD

    if (!dateStr) {
        return new Response(JSON.stringify({ error: "Date parameter required" }), { status: 400 });
    }

    try {
        const { results } = await env.DB.prepare(
            `SELECT a.id, a.user_id, a.status, a.date, a.created_at, u.name 
       FROM attendance a 
       JOIN users u ON a.user_id = u.id 
       WHERE a.date = ? 
       ORDER BY a.created_at DESC`
        )
            .bind(dateStr)
            .all();

        return new Response(JSON.stringify(results), {
            headers: { "Content-Type": "application/json" },
        });
    } catch (e) {
        return new Response(JSON.stringify({ error: e.message }), { status: 500 });
    }
}

export async function onRequestPost(context) {
    const { env, request } = context;

    try {
        const body = await request.json();
        const { user_id, status, date } = body;

        // Validation
        if (!user_id || !status || !date) {
            return new Response(JSON.stringify({ error: "Missing fields" }), { status: 400 });
        }

        // Check for Weekend (Saturday/Sunday)
        // input date format YYYY-MM-DD (e.g. 2024-02-03)
        // new Date('2024-02-03') is parsed as UTC midnight
        const d = new Date(date);
        const day = d.getUTCDay(); // 0=Sunday, 6=Saturday
        if (day === 0 || day === 6) {
            return new Response(JSON.stringify({ error: "Absensi ditutup pada hari libur (Sabtu/Minggu)." }), { status: 400 });
        }

        // Check if duplicate
        const existing = await env.DB.prepare(
            "SELECT id FROM attendance WHERE user_id = ? AND date = ?"
        )
            .bind(user_id, date)
            .first();

        if (existing) {
            return new Response(JSON.stringify({ error: "User sudah absen hari ini" }), { status: 409 });
        }

        // Insert
        await env.DB.prepare(
            "INSERT INTO attendance (user_id, status, date) VALUES (?, ?, ?)"
        )
            .bind(user_id, status, date)
            .run();

        return new Response(JSON.stringify({ success: true }), { status: 201 });

    } catch (e) {
        return new Response(JSON.stringify({ error: e.message }), { status: 500 });
    }
}
