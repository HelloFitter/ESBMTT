SELECT   *
  FROM   (WITH DATA_TMP
                 AS (SELECT   B.CONSUMER_INVOKE_ID, B.PROVIDER_INVOKE_ID
                       FROM         SERVICE_INVOKE A
                                 LEFT JOIN
                                    INTERFACE_INVOKE B
                                 ON A.INVOKE_ID = PROVIDER_INVOKE_ID
                                    AND B.CONSUMER_INVOKE_ID IS NOT NULL
                              LEFT JOIN
                                 SYSTEM C
                              ON A.SYSTEM_ID = C.SYSTEM_ID
                      WHERE       A.SERVICE_ID = '&1'
                              AND A.OPERATION_ID = '&2'
                              AND UPPER (C.SYSTEM_AB) = UPPER ('&4'))
          SELECT   (SELECT   INVOKE_ID
                      FROM   SERVICE_INVOKE M, SYSTEM N
                     WHERE       M.INVOKE_ID = A.CONSUMER_INVOKE_ID
                             AND M.SYSTEM_ID = N.SYSTEM_ID
                             AND N.SYSTEM_CHINESE_NAME LIKE '%&3%')
                      CONSUMER_INVOKE_ID,
                   (SELECT   CASE
                                WHEN N.SYSTEM_AB IN ('BIBPS','PPL') THEN '0'
                                ELSE NVL (GENERATOR_ID, '1')
                             END
                      FROM   SYSTEM N,    SERVICE_INVOKE M
                                       LEFT JOIN
                                          PROTOCOL Q
                                       ON M.PROTOCOL_ID = Q.PROTOCOL_ID
                     WHERE       M.INVOKE_ID = A.CONSUMER_INVOKE_ID
                             AND M.SYSTEM_ID = N.SYSTEM_ID
                             AND N.SYSTEM_CHINESE_NAME LIKE '%&3%')
                      CONSUMER_PROTOCOL_ID,
                   (SELECT   INVOKE_ID
                      FROM   SERVICE_INVOKE
                     WHERE   INVOKE_ID = A.PROVIDER_INVOKE_ID)
                      PROVIDER_INVOKE_ID,
                   (SELECT   NVL (
                                (SELECT   CASE
                                             WHEN A.SYSTEM_AB IN ('BIBPS','PPL')
                                             THEN
                                                '0'
                                             WHEN A.SYSTEM_AB IN ('CCSNEW','SMSP')
                                             THEN
                                                '1'
                                             ELSE
                                                C.GENERATOR_ID
                                          END
                                   FROM         SYSTEM A
                                             LEFT JOIN
                                                SYSTEM_PROTOCOL B
                                             ON A.SYSTEM_ID = B.SYSTEM_ID
                                          LEFT JOIN
                                             PROTOCOL C
                                          ON B.PROTOCOL_ID = C.PROTOCOL_ID
                                  WHERE   A.SYSTEM_AB = '&4'),
                                '1'
                             )
                      FROM      SERVICE_INVOKE M
                             LEFT JOIN
                                PROTOCOL N
                             ON M.PROTOCOL_ID = N.PROTOCOL_ID
                     WHERE   M.INVOKE_ID = A.PROVIDER_INVOKE_ID)
                      PROVIDER_PROTOCOL_ID
            FROM   DATA_TMP A)
 WHERE   CONSUMER_INVOKE_ID IS NOT NULL;

EXIT;