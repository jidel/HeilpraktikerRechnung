# Backend
This file contains the necessary tables for the application. If you already have data in an oldr version of the backend, be sure to apply the changes listed below.

## Changelog
### v1.4
- Table ``tblRechnungen``
  - new field ``RG_DIAGNOSE``
  - removed field ``BHL_ID``
  - converted values of ``RG_NR`` into numbers (by removing the ``-``)
- Table ``tblPatienten``
  - new field ``PAT_ABWEICHENDE_RG``, type: ``boolean``, default value: ``false`` (check all where there is data in ``PAT_RECHNUNG_*``)
  - convert values of ``PAT_RECHNUNG_ANREDE`` to text (``2 = f``and ``1 = m``)
- Remove table ``tblBehandlungen`` and ``tblBehandlungstermine``
  
