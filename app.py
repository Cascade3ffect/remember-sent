# app.py

from aiohttp import web
import aiosqlite
import jinja2
import aiohttp_jinja2

app = web.Application()

aiohttp_jinja2.setup(
    app,
    loader=jinja2.FileSystemLoader("templates")
)

async def hello(request):
    return web.Response(text="Dashboard Running")
    
@aiohttp_jinja2.template("dashboard.html")    
async def dashboard(request):

    async with aiosqlite.connect("email.db") as db:
        
        db.row_factory = aiosqlite.Row
        
        cursor = await db.execute("""
            SELECT *
            FROM conversations
            ORDER BY ActionScore DESC
        """)

        rows = await cursor.fetchall()
        
        print(f"Rows returned: {len(rows)}")

    return {
        "rows": rows
    }

# async def summary(request):

    # async with aiosqlite.connect("email.db") as db:


        # cursor = await db.execute("""
            # SELECT *
            # FROM conversations
            # ORDER BY ActionScore DESC
        # """)

        # rows = await cursor.fetchall()

    # return web.json_response({
        # "open": data[0],
        # "outstanding": data[1],
        # "maxScore": data[2]
    # })
    
app.router.add_static(
    '/static/',
    path='static',
    name='static'
)

app.router.add_get("/hello", hello)

app.router.add_get("/", dashboard)

# app.router.add_get("/api/summary", summary)

web.run_app(
    app, 
    host="127.0.0.1", 
    port=8080)