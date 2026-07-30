import type { APIRoute } from 'astro';
import { env } from 'cloudflare:workers';
import { ensureSchema, exportRows, type InvoiceFilters } from '../../lib/db';
import { rowsToXlsxZip, XLSX_CONTENT_TYPE } from '../../lib/xlsx';

export const GET: APIRoute = async ({ url }) => {
  await ensureSchema(env.DB);
  const p = url.searchParams;
  const ivaParam = p.get('iva');

  const filters: InvoiceFilters = {
    q: p.get('q') || undefined,
    account: p.get('account') || undefined,
    docType: p.get('docType') || undefined,
    ivaRate: ivaParam && Number.isFinite(Number(ivaParam)) ? Number(ivaParam) : undefined,
    from: p.get('from') || undefined,
    to: p.get('to') || undefined,
    hasPdf: p.get('hasPdf') === '1' ? true : p.get('hasPdf') === '0' ? false : undefined,
    sort: p.get('sort') || 'fecha',
    dir: p.get('dir') === 'asc' ? 'asc' : 'desc',
  };

  const rows = await exportRows(env.DB, filters);
  const headers = [
    'fecha_emision', 'doc_type', 'consecutivo', 'clave', 'emisor_nombre', 'emisor_id',
    'emisor_email', 'receptor_nombre', 'receptor_id', 'receptor_email', 'moneda',
    'tipo_cambio', 'condicion_venta', 'iva_rate', 'total_gravado', 'total_exento',
    'total_exonerado', 'total_descuentos', 'total_venta_neta', 'total_impuesto',
    'total_otros_cargos', 'total_comprobante', 'source_account', 'has_pdf',
  ];
  const zip = rowsToXlsxZip(headers, rows as unknown as Record<string, unknown>[]);
  return new Response(zip.body, {
    headers: {
      'content-type': XLSX_CONTENT_TYPE,
      'content-disposition': `attachment; filename="invoices-${new Date().toISOString().slice(0, 10)}.xlsx"`,
    },
  });
};
