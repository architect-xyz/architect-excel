import { useQuery } from '@apollo/client'
import { isLicensed, License, useLicense } from '@architect/lib/provider/ArchitectMode'
import { Client } from 'app/data'
import { graphql } from 'app/data/graphql'
import { MARKET_FRAG } from 'app/data/queries/markets'
import { addMinutes, roundToNearestMinutes } from 'date-fns'
import { ResultOf } from 'gql.tada'
import { useCallback, useEffect, useState } from 'react'

export const __GET_MARKETS_STATS = graphql(`
  query GetMarketSnapshots($latestAtOrBefore: DateTime) {
    marketsSnapshots(latestAtOrBefore: $latestAtOrBefore) {
      __typename
      marketId
      high24h
      lastPrice
      low24h
      volume24h
      open24h
      bidPrice
      askPrice
    }
  }
`)
const FRAG_SNAPSHOT = graphql(`
  fragment Snapshot on MarketSnapshot {
    __typename
    marketId
    high24h
    lastPrice
    low24h
    volume24h
    open24h
    bidPrice
    askPrice
  }
`)

export type MarketSnapshot = NonNullable<
  ResultOf<typeof __GET_MARKETS_STATS>['marketsSnapshots'][number]
>
type MarketId = string
export function readMarket(client: Client, marketId: MarketId) {
  return client.readFragment({
    id: `Market:${marketId}`,
    fragment: MARKET_FRAG,
    fragmentName: 'MarketFields',
  })
}
export function readMarketStats(client: Client, marketId: MarketId) {
  return client.readFragment({
    id: `MarketSnapshot:${marketId}`,
    fragment: FRAG_SNAPSHOT,
  })
}

/**
 * useMarketStats() is called at the root in useDataFetcher() to set up the
 * polling cycle. After this, all consumers only read data from the cache,
 * automatically updating whenever the data changes from the root query
 *
 * @param [isRootQuery=false] this should only be set by `useDataFetcher()`
 * @param [pollInterval=5_000] this is only configurable for tests
 * */
export function useMarketStatsQuery(isRootQuery = false, pollInterval = 5_000) {
  const license = useLicense()
  const hasLicensed = isLicensed(license)
  const [latestAtOrBefore, setLatestAtOrBefore] = useState<Date>()

  const doUpdateLatestAtOrBefore = useCallback(() => {
    if (hasLicensed) return
    // round to nearest minute to prevent each component having different variables
    setLatestAtOrBefore(roundToNearestMinutes(addMinutes(new Date(), -15)))
  }, [hasLicensed])

  useEffect(() => {
    doUpdateLatestAtOrBefore()
  }, [doUpdateLatestAtOrBefore])

  return useQuery(__GET_MARKETS_STATS, {
    fetchPolicy: isRootQuery ? 'network-only' : 'cache-first',
    pollInterval,
    skip: license === License.UNKNOWN,
    variables: {
      latestAtOrBefore: latestAtOrBefore?.toISOString(),
    },
    notifyOnNetworkStatusChange: true,
    onCompleted: doUpdateLatestAtOrBefore,
  })
}

/**
 * useMarketStatsQueryLoading() will return true when the query hasn't fetched for the first time
 * */
export function useMarketStatsQueryLoading() {
  const [loading, setLoading] = useState(true)
  const { loading: queryLoading } = useMarketStatsQuery()

  useEffect(() => {
    if (loading && !queryLoading) {
      setLoading(false)
    }
  }, [loading, queryLoading])

  return loading
}
