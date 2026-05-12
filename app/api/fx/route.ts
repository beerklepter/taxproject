export async function GET(req: Request) {
  const { searchParams } = new URL(req.url)
  const date = searchParams.get('date')

  if (!date) {
    return Response.json({ rate: '' })
  }

  try {
    const res = await fetch(
      `https://api.cnb.cz/cnbapi/exrates/daily?date=${date}&lang=EN`
    )

    const data = await res.json()

    const eur = data.rates?.find(
      (r: any) => r.currencyCode === 'EUR'
    )

    return Response.json({
      rate: eur?.rate ?? ''
    })
  } catch (e) {
    return Response.json({ rate: '' })
  }
}