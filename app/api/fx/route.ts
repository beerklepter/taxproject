const ISO_DATE = /^\d{4}-\d{2}-\d{2}$/

export async function GET(req: Request) {
  const { searchParams } = new URL(req.url)
  const date = searchParams.get('date')

  if (!date || !ISO_DATE.test(date)) {
    return Response.json({ rate: '' })
  }

  try {
    const res = await fetch(
      `https://api.cnb.cz/cnbapi/exrates/daily?date=${encodeURIComponent(date)}&lang=EN`
    )

    if (!res.ok) {
      return Response.json({ rate: '' })
    }

    const data = await res.json()

    const eur = data.rates?.find(
      (r: { currencyCode?: string }) => r.currencyCode === 'EUR'
    )

    return Response.json({
      rate: eur?.rate ?? ''
    })
  } catch {
    return Response.json({ rate: '' })
  }
}