new Vue({
  el: '#app',

  data: {
    isCheckingUser: true,
    isAuthenticated: false,
    userEmail: '',
    userName: '',

    loginForm: {
      username: '',
      password: ''
    },
    isLoggingIn: false,
    loginError: '',

    isSidebarOpen: false,
    showSidebar: false,
    isDesktopView: true,
    activeView: 'dashboard',

    isLoadingFormData: false,
    isSavingPG: false,
    isSaveSuccess: false,

    showErrorModal: false,
    failedTransaction: null,
    failedErrorMessage: '',

    pgForm: {
      fechaRegistro: new Date().toLocaleDateString('en-CA', { timeZone: 'America/Bogota' }),
      fechaPago: new Date().toLocaleDateString('en-CA', { timeZone: 'America/Bogota' }),
      tipo: '',
      grupo: '',
      categoria: '',
      formaPago: '',
      tercero: '',
      valor: '',
      descripcion: ''
    },

    dbTerceros: [],
    dbFormasPago: [],
    dbTablaPG: [],
    dbFlash: [],
    dbLista: [],

    isLoadingList: false,
    isSyncingBackground: false,
    searchQuery: '',

    filters: {
      tipo: '',
      grupo: '',
      categoria: '',
      tercero: '',
      medio: '',
      periodo: 'mes',
      useFechaPago: false
    },

    tesoreriaFilter: 'hoy',
    dashboardFilter: 'ytd',
    dashTableFilterTipo: '',
    dashTableFilterGrupo: '',
    
    tesoreriaSelectedDefaultSources: ['Nequi', 'Daviplata', 'Cash', 'Bbva Libreton #1782'],
    tesoreriaSelectedSources: [],

    trmUSD: 4000,
    isFetchingTRM: false,

    saldosReales: {},
    isSavingAjuste: {},

    // =====================================================
    // PROYECCIÓN / ESTADO BASE
    // =====================================================
    isLoadingProjection: false,
    isSyncingProjectionBackground: false,
    isSavingProjection: false,
    projectionSaveSuccess: false,

    dbProjectionBudget: [],
    dbProjectionExecution: [],

    projectionFilters: {
      periodo: 'mes',
      tipo: '',
      grupo: '',
      categoria: '',
      tercero: '',
      search: ''
    },

    projectionDrafts: {},
    projectionEditingRows: {},
    projectionDeletedRows: {},
    projectionTasks: {},

    projectionCopyModal: {
      show: false,
      targetDate: new Date().toLocaleDateString('en-CA', { timeZone: 'America/Bogota' })
    },

    projectionCopyPreview: [],

    chartPygInstance: null,
    chartPasivosInstance: null,

    inactivityTimeoutMs: 5 * 60 * 1000,
    inactivityTimer: null,
    activityListenersBound: false

  },

  computed: {

    hayFiltroTesoreria() {
      return this.tesoreriaSelectedSources.length !== this.fuentesSeleccionablesTesoreria.length;
    },
    displayName() {
      return this.userName || this.userEmail || 'Usuario';
    },

    userInitials() {
      return (this.userName || this.userEmail || 'U').charAt(0).toUpperCase();
    },

    isLoginFormValid() {
      return this.loginForm.username.trim() !== '' && this.loginForm.password.trim() !== '';
    },

    isPGFormValid() {
      return (
        this.pgForm.fechaRegistro !== '' &&
        this.pgForm.fechaPago !== '' &&
        this.pgForm.tipo !== '' &&
        this.pgForm.grupo !== '' &&
        this.pgForm.categoria !== '' &&
        this.pgForm.formaPago !== '' &&
        this.pgForm.tercero !== '' &&
        this.pgForm.valor !== '' &&
        this.pgForm.descripcion.trim() !== ''
      );
    },

    formasPagoOptionsActivas() {
      return this.dbFormasPago
        .filter(item => item.estado_fuente !== 'n')
        .map(item => item.medio)
        .sort();
    },

    formasPagoOptions() {
      return this.dbFormasPago.map(item => item.medio).sort();
    },

    tercerosOptions() {
      return [...this.dbTerceros].sort();
    },

    tiposOptions() {
      return [...new Set(this.dbTablaPG.map(item => item.tipo))].sort();
    },

    categoriasOptions() {
      if (!this.pgForm.tipo) return [];
      return [...new Set(
        this.dbTablaPG
          .filter(item => item.tipo === this.pgForm.tipo)
          .map(item => item.categoria)
      )].sort();
    },

    gruposOptions() {
      if (!this.pgForm.tipo || !this.pgForm.categoria) return [];
      return [...new Set(
        this.dbTablaPG
          .filter(item => item.tipo === this.pgForm.tipo && item.categoria === this.pgForm.categoria)
          .map(item => item.grupo)
      )].sort();
    },

    listGruposOptions() {
      return [...new Set(this.dbTablaPG.map(item => item.grupo))].sort();
    },

    listCategoriasOptions() {
      return [...new Set(this.dbTablaPG.map(item => item.categoria))].sort();
    },

    totalIngresos() {
      return this.dbListaFiltrada
        .filter(t => t.tipo === 'Ingreso')
        .reduce((sum, t) => sum + (Number(t.valor) || 0), 0);
    },

    totalGastos() {
      return this.dbListaFiltrada
        .filter(t => t.tipo === 'Gasto')
        .reduce((sum, t) => sum + (Number(t.valor) || 0), 0);
    },

    totalActivos() {
      return this.dbListaFiltrada
        .filter(t => t.tipo === 'Activo')
        .reduce((sum, t) => sum + (Number(t.valor) || 0), 0);
    },

    totalPasivos() {
      return this.dbListaFiltrada
        .filter(t => t.tipo === 'Pasivo')
        .reduce((sum, t) => sum + (Number(t.valor) || 0), 0);
    },

    utilidadPorcentaje() {
      if (this.totalIngresos === 0) return '-%';
      const u = ((this.totalIngresos - this.totalGastos) / this.totalIngresos) * 100;
      return u <= 0 ? '-%' : u.toFixed(1) + '%';
    },

    todayFormatted() {
      const d = new Date();
      return d.toLocaleDateString('es-CO', {
        day: '2-digit',
        month: '2-digit',
        year: 'numeric',
        timeZone: 'America/Bogota'
      });
    },

    saldosAgrupados() {
      const mediosElegibles = this.dbFormasPago.filter(f => {
        const tf = (f.tipo_fuente || '').toLowerCase();
        return (
          tf.includes('efectivo') ||
          tf.includes('aplicacion') ||
          tf.includes('tarjeta debito') ||
          tf.includes('tarjeta débito')
        );
      });

      const today = new Date();
      let cutoffDate = new Date();

      if (this.tesoreriaFilter === 'mes_anterior') {
        cutoffDate = new Date(today.getFullYear(), today.getMonth(), 0);
      } else if (this.tesoreriaFilter === 'año_anterior') {
        cutoffDate = new Date(today.getFullYear() - 1, 11, 31);
      }

      cutoffDate.setHours(23, 59, 59, 999);

      const saldosIndividuales = mediosElegibles.map(medio => {
        const movsMedio = this.dbLista.filter(m => {
          const pDate = new Date((m.fecha_payment || m.fecha_pago) + 'T00:00:00');
          return m.medio === medio.medio && pDate <= cutoffDate;
        });

        const saldoFinal = movsMedio.reduce((acc, mov) => {
          let val = Number(mov.valor) || 0;
          const tipo = (mov.tipo || '').trim();
          const desc = (mov.descripcion || '').trim().toLowerCase();

          if (tipo === 'Movimiento') return acc + val;
          if ((tipo === 'Activo' || tipo === 'Pasivo') && desc === 'apertura') val = 0;
          if (tipo === 'Gasto' || tipo === 'Activo') return acc - val;
          return acc + val;
        }, 0);

        return {
          nombre: medio.medio,
          tipo_fuente: medio.tipo_fuente,
          saldo: saldoFinal,
          estado_fuente: medio.estado_fuente || 'h',
          moneda: medio.moneda || 'COP'
        };
      });

      const grupos = {};
      saldosIndividuales.forEach(s => {
        if (!grupos[s.tipo_fuente]) {
          grupos[s.tipo_fuente] = { tipo: s.tipo_fuente, medios: [], total: 0 };
        }
        grupos[s.tipo_fuente].medios.push(s);

        const saldoCOP = s.moneda === 'USD' ? s.saldo * (this.trmUSD || 0) : s.saldo;
        grupos[s.tipo_fuente].total += saldoCOP;
      });

      Object.values(grupos).forEach(g => {
        g.medios.sort((a, b) => {
          if (a.estado_fuente === 'n' && b.estado_fuente !== 'n') return 1;
          if (a.estado_fuente !== 'n' && b.estado_fuente === 'n') return -1;
          return b.saldo - a.saldo;
        });
      });

      return Object.values(grupos).sort((a, b) => b.total - a.total);
    },

    saldosPlanos() {
      let list = [];
      this.saldosAgrupados.forEach(g => { list = list.concat(g.medios); });
      return list;
    },

    fuentesSeleccionablesTesoreria() {
      return this.saldosPlanos
        .filter(medio => medio.estado_fuente === 'h')
        .sort((a, b) => a.nombre.localeCompare(b.nombre));
    },

    totalSeleccionadoTesoreria() {
      return this.fuentesSeleccionablesTesoreria.reduce((acc, medio) => {
        if (!this.tesoreriaSelectedSources.includes(medio.nombre)) return acc;

        const saldoCOP = medio.moneda === 'USD'
          ? medio.saldo * (this.trmUSD || 0)
          : medio.saldo;

        return acc + saldoCOP;
      }, 0);
    },

    cantidadFuentesSeleccionadasTesoreria() {
      return this.fuentesSeleccionablesTesoreria.filter(medio =>
        this.tesoreriaSelectedSources.includes(medio.nombre)
      ).length;
    },

    totalConsolidadoTesoreria() {
      return this.saldosAgrupados.reduce((acc, g) => acc + g.total, 0);
    },

    totalCuadreSistema() {
      return this.saldosPlanos.reduce((acc, m) => {
        const valCOP = m.moneda === 'USD' ? m.saldo * (this.trmUSD || 0) : m.saldo;
        return acc + valCOP;
      }, 0);
    },

    totalCuadreDiferencia() {
      return this.saldosPlanos.reduce((acc, m) => {
        const diff = this.getDiff(m);
        const diffCOP = m.moneda === 'USD' ? diff * (this.trmUSD || 0) : diff;
        return acc + diffCOP;
      }, 0);
    },

    totalCuadreReal() {
      return this.saldosPlanos.reduce((acc, m) => {
        let real = this.saldosReales[m.nombre];
        let val = (real === undefined || real === '') ? m.saldo : this.parseToNumber(real);
        let valCOP = m.moneda === 'USD' ? val * (this.trmUSD || 0) : val;
        return acc + valCOP;
      }, 0);
    },

    dashboardDates() {
      const today = new Date();
      const y = today.getFullYear();
      const m = today.getMonth();
      const d = today.getDate();

      let currStart = new Date(), currEnd = new Date(), prevStart = new Date(), prevEnd = new Date();
      let compareText = '';

      switch (this.dashboardFilter) {
        case 'este_mes':
          currStart = new Date(y, m, 1);
          currEnd = new Date(y, m, d);
          prevStart = new Date(y, m - 1, 1);
          prevEnd = new Date(y, m, 0);
          compareText = 'vs. Mes Anterior';
          break;
        case 'mes_pasado':
          currStart = new Date(y, m - 1, 1);
          currEnd = new Date(y, m, 0);
          prevStart = new Date(y - 1, m - 1, 1);
          prevEnd = new Date(y - 1, m, 0);
          compareText = 'vs. Mismo Mes Año Ant.';
          break;
        case 'ytd':
          currStart = new Date(y, 0, 1);
          currEnd = new Date(y, m, d);
          prevStart = new Date(y - 1, 0, 1);
          prevEnd = new Date(y - 1, m, d);
          compareText = 'vs. Mismo Periodo Año Ant.';
          break;
        case 'año_pasado':
          currStart = new Date(y - 1, 0, 1);
          currEnd = new Date(y - 1, 11, 31);
          prevStart = new Date(y - 2, 0, 1);
          prevEnd = new Date(y - 2, 11, 31);
          compareText = 'vs. Año Ant. del Pasado';
          break;
      }

      currStart.setHours(0, 0, 0, 0);
      currEnd.setHours(23, 59, 59, 999);
      prevStart.setHours(0, 0, 0, 0);
      prevEnd.setHours(23, 59, 59, 999);

      return { currStart, currEnd, prevStart, prevEnd, compareText };
    },

    dashboardData() {
      const { currStart, currEnd, prevStart, prevEnd, compareText } = this.dashboardDates;

      let ingCurr = 0, ingPrev = 0, gasCurr = 0, gasPrev = 0, actCurr = 0, actPrev = 0, pasCurr = 0, pasPrev = 0;
      let dsrValCurr = 0, dsrValPrev = 0;

      this.dbLista.forEach(t => {
        if (t.estado === 'n') return;
        const val = Number(t.valor) || 0;
        const fStr = t.fecha_registration || t.fecha_registro;
        if (!fStr) return;

        const itemDate = new Date(fStr + 'T00:00:00');

        if (t.tipo === 'Ingreso') {
          if (itemDate >= currStart && itemDate <= currEnd) ingCurr += val;
          if (itemDate >= prevStart && itemDate <= prevEnd) ingPrev += val;
        } else if (t.tipo === 'Gasto') {
          if (itemDate >= currStart && itemDate <= currEnd) gasCurr += val;
          if (itemDate >= prevStart && itemDate <= prevEnd) gasPrev += val;
        }

        if (t.tipo === 'Activo') {
          const grp = (t.grupo || '').trim().toLowerCase();
          if (grp.includes('ahorro') || grp.includes('inversio')) {
            if (itemDate <= currEnd) actCurr += val;
            if (itemDate <= prevEnd) actPrev += val;
          }
        } else if (t.tipo === 'Pasivo') {
          if (itemDate <= currEnd) pasCurr += val;
          if (itemDate <= prevEnd) pasPrev += val;

          if (val < 0) {
            if (itemDate >= currStart && itemDate <= currEnd) dsrValCurr += Math.abs(val);
            if (itemDate >= prevStart && itemDate <= prevEnd) dsrValPrev += Math.abs(val);
          }
        }
      });

      let utiCurr = ingCurr - gasCurr;
      let utiPrev = ingPrev - gasPrev;
      let utiPctCurr = ingCurr ? (utiCurr / ingCurr * 100) : 0;

      let dsrCurr = ingCurr ? (dsrValCurr / ingCurr * 100) : 0;
      let dsrPrev = ingPrev ? (dsrValPrev / ingPrev * 100) : 0;

      const calcDiff = (curr, prev) => {
        const diff = curr - prev;
        let pct = 0;
        if (prev !== 0) pct = (diff / Math.abs(prev)) * 100;
        else if (curr > 0) pct = 100;
        else if (curr < 0) pct = -100;
        return { curr, prev, diff, pct };
      };

      return {
        compareText,
        ingresos: calcDiff(ingCurr, ingPrev),
        gastos: calcDiff(gasCurr, gasPrev),
        activos: calcDiff(actCurr, actPrev),
        pasivos: calcDiff(pasCurr, pasPrev),
        utilidad: { ...calcDiff(utiCurr, utiPrev), margin: utiPctCurr },
        dsr: { curr: dsrCurr, diff: dsrCurr - dsrPrev }
      };
    },

    dashboardTables() {
      const { currStart, currEnd, prevStart, prevEnd } = this.dashboardDates;

      let mapGastos = {}, mapIngresos = {}, mapActivos = {}, mapPasivos = {};
      let totalTipo = { Ingreso: 0, Gasto: 0, Activo: 0, Pasivo: 0 };

      this.dbLista.forEach(t => {
        if (t.estado === 'n') return;
        const val = Number(t.valor) || 0;
        const fStr = t.fecha_registration || t.fecha_registro;
        if (!fStr) return;
        const itemDate = new Date(fStr + 'T00:00:00');

        const processMap = (map, key, isFlow, tipo) => {
          if (!key) return;
          if (!map[key]) map[key] = { curr: 0, prev: 0, sum_period: 0, activeMonths: new Set() };

          if (isFlow) {
            if (itemDate >= currStart && itemDate <= currEnd) {
              map[key].curr += val;
              map[key].sum_period += val;
              map[key].activeMonths.add(itemDate.getFullYear() + '-' + itemDate.getMonth());
              totalTipo[tipo] += val;
            }
            if (itemDate >= prevStart && itemDate <= prevEnd) map[key].prev += val;
          } else {
            if (itemDate <= currEnd) map[key].curr += val;
            if (itemDate <= prevEnd) map[key].prev += val;
            if (itemDate >= currStart && itemDate <= currEnd) {
              map[key].sum_period += val;
              map[key].activeMonths.add(itemDate.getFullYear() + '-' + itemDate.getMonth());
            }
          }
        };

        if (t.tipo === 'Gasto') {
          processMap(mapGastos, t.grupo, true, 'Gasto');
        } else if (t.tipo === 'Ingreso') {
          processMap(mapIngresos, t.categoria, true, 'Ingreso');
        } else if (t.tipo === 'Pasivo') {
          processMap(mapPasivos, `${t.categoria} • ${t.tercero}`, false, 'Pasivo');
        } else if (t.tipo === 'Activo') {
          const grp = (t.grupo || '').trim().toLowerCase();
          if (grp.includes('ahorro') || grp.includes('inversio')) {
            processMap(mapActivos, `${t.categoria} • ${t.tercero}`, false, 'Activo');
          }
        }
      });

      totalTipo['Activo'] = Object.values(mapActivos).reduce((s, x) => s + x.curr, 0);
      totalTipo['Pasivo'] = Object.values(mapPasivos).reduce((s, x) => s + x.curr, 0);

      const formatMap = (map, tipo) => {
        return Object.keys(map).map(k => {
          const c = map[k].curr;
          const p = map[k].prev;
          const d = c - p;
          let pt = 0;

          if (p !== 0) pt = (d / Math.abs(p)) * 100;
          else if (c > 0) pt = 100;
          else if (c < 0) pt = -100;

          let pctTotal = 0;
          if (totalTipo[tipo] !== 0) {
            pctTotal = (Math.abs(c) / Math.abs(totalTipo[tipo])) * 100;
          }

          let activeMonthsCount = map[k].activeMonths.size > 0 ? map[k].activeMonths.size : 1;
          let promedio = map[k].sum_period / activeMonthsCount;

          return { label: k, curr: c, prev: p, diff: d, pct: pt, pctTotal, promedio };
        })
        .filter(item => item.curr !== 0)
        .sort((a, b) => Math.abs(b.curr) - Math.abs(a.curr));
      };

      return {
        gastos: formatMap(mapGastos, 'Gasto'),
        ingresos: formatMap(mapIngresos, 'Ingreso'),
        activos: formatMap(mapActivos, 'Activo'),
        pasivos: formatMap(mapPasivos, 'Pasivo')
      };
    },

    dashTableGruposOptions() {
      let list = this.dbLista;
      if (this.dashTableFilterTipo) {
        list = list.filter(t => t.tipo === this.dashTableFilterTipo);
      }
      return [...new Set(list.map(t => t.grupo))].filter(Boolean).sort();
    },

    dashboardDetailedTable() {
      const { currStart, currEnd, prevStart, prevEnd } = this.dashboardDates;
      let map = {};
      let totalPorTipo = { Ingreso: 0, Gasto: 0, Activo: 0, Pasivo: 0 };

      this.dbLista.forEach(t => {
        if (t.estado === 'n') return;
        const val = Number(t.valor) || 0;
        const fStr = t.fecha_registration || t.fecha_registro;
        if (!fStr) return;
        const itemDate = new Date(fStr + 'T00:00:00');

        const key = `${t.tipo}|${t.grupo}|${t.categoria}`;

        if (!map[key]) {
          map[key] = {
            tipo: t.tipo,
            grupo: t.grupo,
            categoria: t.categoria,
            curr: 0,
            prev: 0,
            count: 0,
            sum_period: 0,
            activeMonths: new Set()
          };
        }

        let isFlow = (t.tipo === 'Ingreso' || t.tipo === 'Gasto');

        if (isFlow) {
          if (itemDate >= currStart && itemDate <= currEnd) {
            map[key].curr += val;
            map[key].count += 1;
            map[key].sum_period += val;
            map[key].activeMonths.add(itemDate.getFullYear() + '-' + itemDate.getMonth());
            totalPorTipo[t.tipo] += val;
          }
          if (itemDate >= prevStart && itemDate <= prevEnd) {
            map[key].prev += val;
          }
        } else {
          if (itemDate <= currEnd) map[key].curr += val;
          if (itemDate <= prevEnd) map[key].prev += val;

          if (itemDate >= currStart && itemDate <= currEnd) {
            map[key].count += 1;
            map[key].sum_period += val;
            map[key].activeMonths.add(itemDate.getFullYear() + '-' + itemDate.getMonth());
          }
        }
      });

      totalPorTipo['Activo'] = Object.values(map).filter(x => x.tipo === 'Activo').reduce((s, x) => s + x.curr, 0);
      totalPorTipo['Pasivo'] = Object.values(map).filter(x => x.tipo === 'Pasivo').reduce((s, x) => s + x.curr, 0);

      let results = Object.values(map).map(row => {
        const d = row.curr - row.prev;
        let pt = 0;

        if (row.prev !== 0) pt = (d / Math.abs(row.prev)) * 100;
        else if (row.curr > 0) pt = 100;
        else if (row.curr < 0) pt = -100;

        let pctTotal = 0;
        if (totalPorTipo[row.tipo] !== 0) {
          pctTotal = (Math.abs(row.curr) / Math.abs(totalPorTipo[row.tipo])) * 100;
        }

        let activeMonthsCount = row.activeMonths.size > 0 ? row.activeMonths.size : 1;
        let promedio = row.sum_period / activeMonthsCount;

        return { ...row, diff: d, pct: pt, pctTotal, promedio };
      }).filter(r => r.curr !== 0 || r.prev !== 0);

      if (this.dashTableFilterTipo) {
        results = results.filter(r => r.tipo === this.dashTableFilterTipo);
      }

      if (this.dashTableFilterGrupo) {
        results = results.filter(r => r.grupo === this.dashTableFilterGrupo);
      }

      results.sort((a, b) => {
        if (a.tipo !== b.tipo) return a.tipo.localeCompare(b.tipo);
        return Math.abs(b.curr) - Math.abs(a.curr);
      });

      return results;
    },

    dashboardChartsData() {
      const today = new Date();
      let monthsInfo = [];

      for (let i = 11; i >= 0; i--) {
        let mDate = new Date(today.getFullYear(), today.getMonth() - i, 1);
        monthsInfo.push({
          label: mDate.toLocaleDateString('es-CO', { month: 'short', year: '2-digit' }).replace('.', ''),
          y: mDate.getFullYear(),
          m: mDate.getMonth(),
          start: new Date(mDate.getFullYear(), mDate.getMonth(), 1),
          end: new Date(mDate.getFullYear(), mDate.getMonth() + 1, 0, 23, 59, 59)
        });
      }

      let arrIng = new Array(12).fill(0);
      let arrGas = new Array(12).fill(0);
      let arrUtiPct = new Array(12).fill(0);
      let arrPasMov = new Array(12).fill(0);
      let arrPasSaldo = new Array(12).fill(0);

      let pasivosAcumuladosAntes = 0;

      this.dbLista.forEach(t => {
        if (t.estado === 'n') return;
        const val = Number(t.valor) || 0;
        const fStr = t.fecha_registration || t.fecha_registro;
        if (!fStr) return;
        const itemDate = new Date(fStr + 'T00:00:00');

        if (t.tipo === 'Pasivo' && itemDate < monthsInfo[0].start) {
          pasivosAcumuladosAntes += val;
        }

        for (let i = 0; i < 12; i++) {
          let mi = monthsInfo[i];
          if (itemDate >= mi.start && itemDate <= mi.end) {
            if (t.tipo === 'Ingreso') arrIng[i] += val;
            else if (t.tipo === 'Gasto') arrGas[i] += val;
            else if (t.tipo === 'Pasivo') arrPasMov[i] += val;
            break;
          }
        }
      });

      let saldoCorrientePasivo = pasivosAcumuladosAntes;
      for (let i = 0; i < 12; i++) {
        let uti = arrIng[i] - arrGas[i];
        arrUtiPct[i] = arrIng[i] ? (uti / arrIng[i] * 100) : 0;

        saldoCorrientePasivo += arrPasMov[i];
        arrPasSaldo[i] = saldoCorrientePasivo;
      }

      return {
        labels: monthsInfo.map(m => m.label.toUpperCase()),
        ingresos: arrIng,
        gastos: arrGas,
        utilidadPct: arrUtiPct,
        pasivosMov: arrPasMov,
        pasivosSaldo: arrPasSaldo
      };
    },

    dbListaFiltrada() {
      return this.dbLista.filter(t => {
        const refDateStr = this.filters.useFechaPago
          ? (t.fecha_payment || t.fecha_pago || '')
          : (t.fecha_registration || t.fecha_registro || '');

        if (!refDateStr && this.filters.periodo) return false;

        const itemDate = new Date(refDateStr + 'T00:00:00');
        const today = new Date();
        let matchPeriodo = true;

        if (this.filters.periodo) {
          const y = itemDate.getFullYear();
          const m = itemDate.getMonth();
          const currentY = today.getFullYear();
          const currentM = today.getMonth();

          switch (this.filters.periodo) {
            case 'año':
              matchPeriodo = (y === currentY);
              break;
            case 'año_anterior':
              matchPeriodo = (y === currentY - 1);
              break;
            case 'mes':
              matchPeriodo = (y === currentY && m === currentM);
              break;
            case 'mes_anterior': {
              const prev = new Date(today.getFullYear(), today.getMonth() - 1, 1);
              matchPeriodo = (y === prev.getFullYear() && m === prev.getMonth());
              break;
            }
            case 'semana': {
              const day = today.getDay();
              const diffToMonday = today.getDate() - day + (day === 0 ? -6 : 1);
              const startOfWeek = new Date(today.getFullYear(), today.getMonth(), diffToMonday);
              startOfWeek.setHours(0, 0, 0, 0);
              const endOfWeek = new Date(startOfWeek);
              endOfWeek.setDate(startOfWeek.getDate() + 6);
              endOfWeek.setHours(23, 59, 59, 999);
              matchPeriodo = (itemDate >= startOfWeek && itemDate <= endOfWeek);
              break;
            }
          }
        }

        if (!matchPeriodo) return false;

        const s = this.searchQuery.toLowerCase();
        const mSearch = !s ||
          (t.categoria || '').toLowerCase().includes(s) ||
          (t.tercero || '').toLowerCase().includes(s) ||
          (t.grupo || '').toLowerCase().includes(s) ||
          (t.descripcion || '').toLowerCase().includes(s);

        const mTipo = !this.filters.tipo || t.tipo === this.filters.tipo;
        const mGrupo = !this.filters.grupo || t.grupo === this.filters.grupo;
        const mCat = !this.filters.categoria || t.categoria === this.filters.categoria;
        const mTer = !this.filters.tercero || t.tercero === this.filters.tercero;
        const mMedio = !this.filters.medio || t.medio === this.filters.medio;

        return mSearch && mTipo && mGrupo && mCat && mTer && mMedio;
      });
    },
    
    
    // =====================================================
    // PROYECCIÓN / OPCIONES DE FILTRO
    // =====================================================
    projectionTiposOptions() {
      return [...new Set(this.dbTablaPG.map(item => item.tipo))].filter(Boolean).sort();
    },

    projectionGruposFilterOptions() {
      let base = this.dbTablaPG;

      if (this.projectionFilters.tipo) {
        base = base.filter(item => item.tipo === this.projectionFilters.tipo);
      }

      return [...new Set(base.map(item => item.grupo))].filter(Boolean).sort();
    },

    projectionCategoriasFilterOptions() {
      let base = this.dbTablaPG;

      if (this.projectionFilters.tipo) {
        base = base.filter(item => item.tipo === this.projectionFilters.tipo);
      }

      if (this.projectionFilters.grupo) {
        base = base.filter(item => item.grupo === this.projectionFilters.grupo);
      }

      return [...new Set(base.map(item => item.categoria))].filter(Boolean).sort();
    },

    // =====================================================
    // PROYECCIÓN / PERIODO ACTUAL
    // =====================================================
    projectionCurrentPeriodLabel() {
      const map = {
        mes: 'Mes actual',
        mes_siguiente: 'Mes siguiente',
        mes_anterior: 'Mes anterior',
        ytd: 'Año a la fecha',
        año: 'Año actual',
        año_anterior: 'Año anterior',
        '': 'Todo'
      };

      return map[this.projectionFilters.periodo] || 'Periodo';
    },

    // =====================================================
    // PROYECCIÓN / PRESUPUESTO EFECTIVO
    // Mezcla BD real + tareas pendientes para simular en línea
    // =====================================================
    projectionEffectiveBudgetRows() {
      let rows = this.dbProjectionBudget
        .filter(item => (item.estado || 'h') !== 'n')
        .map(item => ({ ...item }));

      const tasks = Object.values(this.projectionTasks || {});

      tasks.forEach(task => {
        if (task.action === 'create') {
          rows.push({
            idx: '',
            id: '',
            _pending: true,
            _pendingCreate: true,
            _rowKey: task.rowKey,
            ...(task.record || {})
          });
        }

        if (task.action === 'update') {
          const index = rows.findIndex(item =>
            String(item.idx || item.id) === String(task.id)
          );

          if (index !== -1) {
            rows.splice(index, 1, {
              ...rows[index],
              ...(task.record || {}),
              _pending: true,
              _rowKey: task.rowKey
            });
          }
        }

        if (task.action === 'delete') {
          const index = rows.findIndex(item =>
            String(item.idx || item.id) === String(task.id)
          );

          if (index !== -1) {
            rows.splice(index, 1, {
              ...rows[index],
              _pendingDelete: true,
              _rowKey: task.rowKey
            });
          }
        }
      });

      return rows;
    },

    // =====================================================
    // PROYECCIÓN / TABLA COMPARATIVA FLUJO DE CAJA
    // Match por: tipo + grupo + categoria + tercero + flujo
    // =====================================================
    projectionTableRows() {
      const range = this.getProjectionDateRange(this.projectionFilters.periodo);

      const budgetRows = this.projectionEffectiveBudgetRows.filter(item => {
        if ((item.estado || 'h') === 'n') return false;
        return this.isProjectionDateInRange(item.fecha_ejecucion, range);
      });

      const executionRows = this.dbProjectionExecution.filter(item => {
        if ((item.estado || 'h') === 'n') return false;
        const f = item.fecha_registro || item.fecha_registration || item.fecha_pago || item.fecha_payment;
        return this.isProjectionDateInRange(f, range);
      });

      const rows = [];
      const budgetByFull = {};
      const budgetByCategory = {};
      const budgetByGroup = {};
      const hasBudgetByCategory = {};

      budgetRows.forEach(item => {
        const rawValue = Number(item.valor) || 0;
        const flujo = this.getProjectionFlowDirection(item.tipo, rawValue);
        const key = item._rowKey || ('b-' + (item.idx || item.id));

        const row = {
          key: key,
          budgetId: item._pendingCreate ? null : (item.idx || item.id),
          isVirtual: false,
          isPendingCreate: !!item._pendingCreate,

          fecha_ejecucion: item.fecha_ejecucion || '',
          tipo: item.tipo || '',
          flujo: flujo,
          grupo: item.grupo || '',
          categoria: item.categoria || '',
          tercero: item.tercero || '',
          descripcion: item.descripcion || '',

          valor_original: rawValue,
          presupuesto: this.getProjectionComparableValue(item.tipo, rawValue),
          ejecutado: 0,
          detalle: []
        };

        rows.push(row);

        // Si está eliminado en tareas pendientes, se ve en tabla,
        // pero NO participa como presupuesto activo para asignar ejecución.
        if (item._pendingDelete) {
          return;
        }

        const groupKey = this.buildProjectionMatchKey(row.tipo, row.grupo, '', '', row.flujo);
        const catKey = this.buildProjectionMatchKey(row.tipo, row.grupo, row.categoria, '', row.flujo);
        const fullKey = this.buildProjectionMatchKey(row.tipo, row.grupo, row.categoria, row.tercero, row.flujo);

        if (row.categoria) {
          hasBudgetByCategory[catKey] = true;
        }

        if (row.categoria && row.tercero) {
          budgetByFull[fullKey] = row;
        } else if (row.categoria) {
          budgetByCategory[catKey] = row;
        } else {
          budgetByGroup[groupKey] = row;
        }
      });

      const virtualRows = {};

      executionRows.forEach(item => {
        const tipo = item.tipo || '';
        const grupo = item.grupo || '';
        const categoria = item.categoria || '';
        const tercero = item.tercero || '';
        const rawValue = Number(item.valor) || 0;

        const flujo = this.getProjectionFlowDirection(tipo, rawValue);
        const valor = this.getProjectionComparableValue(tipo, rawValue);

        const fullKey = this.buildProjectionMatchKey(tipo, grupo, categoria, tercero, flujo);
        const catKey = this.buildProjectionMatchKey(tipo, grupo, categoria, '', flujo);
        const groupKey = this.buildProjectionMatchKey(tipo, grupo, '', '', flujo);

        let target = null;

        if (budgetByFull[fullKey]) {
          target = budgetByFull[fullKey];
        } else if (budgetByCategory[catKey]) {
          target = budgetByCategory[catKey];
        } else if (budgetByGroup[groupKey]) {
          target = budgetByGroup[groupKey];
        }

        if (target) {
          target.ejecutado += valor;
          target.detalle.push(item);
          return;
        }

        // Para ejecución real sin presupuesto, desglosamos mínimo hasta categoría.
        // No bajamos a tercero para evitar demasiada granularidad.
        const virtualCat = categoria || 'Sin categoría';
        const virtualTer = '';

        const virtualKey = 'u-' + this.buildProjectionMatchKey(tipo, grupo, virtualCat, virtualTer, flujo);

        if (!virtualRows[virtualKey]) {
          virtualRows[virtualKey] = {
            key: virtualKey,
            budgetId: null,
            isVirtual: true,

            fecha_ejecucion: '',
            tipo: tipo,
            flujo: flujo,
            grupo: grupo,
            categoria: virtualCat,
            tercero: virtualTer,
            descripcion: 'Ejecución sin presupuesto',

            valor_original: 0,
            presupuesto: 0,
            ejecutado: 0,
            detalle: []
          };
        }

        virtualRows[virtualKey].ejecutado += valor;
        virtualRows[virtualKey].detalle.push(item);
      });

      Object.values(virtualRows).forEach(row => {
        rows.push(row);
      });

      return rows.map(row => {
        const porcentaje = row.presupuesto > 0
          ? (row.ejecutado / row.presupuesto) * 100
          : (row.ejecutado > 0 ? 999 : 0);

        const diferencia = row.ejecutado - row.presupuesto;

        return {
          ...row,
          porcentaje: porcentaje,
          porcentajeLabel: row.presupuesto > 0
            ? porcentaje.toFixed(1) + '%'
            : (row.ejecutado > 0 ? 'Sin ppto' : '0.0%'),
          diferencia: diferencia,
          alerta: this.getProjectionAlert(row.flujo, porcentaje, row.presupuesto, row.ejecutado)
        };
      }).sort((a, b) => {
        // Las filas nuevas pendientes siempre arriba
        if (a.isPendingCreate && !b.isPendingCreate) return -1;
        if (!a.isPendingCreate && b.isPendingCreate) return 1;

        // Si ambas son nuevas, la más reciente primero
        if (a.isPendingCreate && b.isPendingCreate) {
          const aTime = Number(String(a.key).replace('new-', '')) || 0;
          const bTime = Number(String(b.key).replace('new-', '')) || 0;
          return bTime - aTime;
        }

        // Luego mantenemos el orden financiero normal
        if (a.flujo !== b.flujo) {
          return a.flujo === 'salida' ? -1 : 1;
        }

        if (a.tipo !== b.tipo) return a.tipo.localeCompare(b.tipo);
        if (a.grupo !== b.grupo) return a.grupo.localeCompare(b.grupo);

        return Math.abs(b.ejecutado) - Math.abs(a.ejecutado);
      });
    },

    // =====================================================
    // PROYECCIÓN / TABLA FILTRADA
    // =====================================================
    projectionTableRowsFiltered() {
      return this.projectionTableRows.filter(row => {
        const f = this.projectionFilters;

        if (f.tipo && row.tipo !== f.tipo) return false;
        if (f.grupo && row.grupo !== f.grupo) return false;
        if (f.categoria && row.categoria !== f.categoria) return false;
        if (f.tercero && row.tercero !== f.tercero) return false;

        const s = (f.search || '').toLowerCase().trim();

        if (s) {
          const text = [
            row.tipo,
            row.grupo,
            row.categoria,
            row.tercero,
            row.descripcion
          ].join(' ').toLowerCase();

          if (!text.includes(s)) return false;
        }

        return true;
      });
    },

    // =====================================================
    // PROYECCIÓN / TOTALES FLUJO DE CAJA
    // =====================================================
    projectionTotals() {
      const totals = this.projectionTableRowsFiltered.reduce((acc, row) => {
        if (this.projectionDeletedRows[row.key]) return acc;

        const presupuesto = Number(row.presupuesto) || 0;
        const ejecutado = Number(row.ejecutado) || 0;

        if (row.flujo === 'entrada') {
          acc.presupuestoEntradas += presupuesto;
          acc.ejecutadoEntradas += ejecutado;
        } else {
          acc.presupuestoSalidas += presupuesto;
          acc.ejecutadoSalidas += ejecutado;
        }

        if (row.isVirtual && row.flujo === 'salida') {
          acc.noPresupuestadoSalida += ejecutado;
          acc.noPresupuestadoSalidaItems += 1;
        }

        if (row.isVirtual && row.flujo === 'entrada') {
          acc.noPresupuestadoEntrada += ejecutado;
          acc.noPresupuestadoEntradaItems += 1;
        }

        if (row.alerta && row.alerta.level === 'danger') {
          acc.alertas += 1;
        }

        return acc;
      }, {
        presupuestoEntradas: 0,
        ejecutadoEntradas: 0,

        presupuestoSalidas: 0,
        ejecutadoSalidas: 0,

        noPresupuestadoSalida: 0,
        noPresupuestadoSalidaItems: 0,

        noPresupuestadoEntrada: 0,
        noPresupuestadoEntradaItems: 0,

        alertas: 0
      });

      totals.netoPresupuesto = totals.presupuestoEntradas - totals.presupuestoSalidas;
      totals.netoEjecutado = totals.ejecutadoEntradas - totals.ejecutadoSalidas;
      totals.diferenciaNeto = totals.netoEjecutado - totals.netoPresupuesto;

      // Alias de compatibilidad por si algo del HTML anterior todavía los usa
      totals.presupuesto = totals.presupuestoSalidas;
      totals.ejecutado = totals.ejecutadoSalidas;
      totals.diferencia = totals.ejecutadoSalidas - totals.presupuestoSalidas;
      totals.noPresupuestado = totals.noPresupuestadoSalida;
      totals.noPresupuestadoItems = totals.noPresupuestadoSalidaItems;

      return totals;
    },

    // =====================================================
    // PROYECCIÓN / MÉTRICAS KPI
    // Ratios y diferencias para tarjetas superiores
    // =====================================================
    projectionKpiMetrics() {
      const t = this.projectionTotals;

      const pctAbs = function(real, presupuesto) {
        const r = Math.abs(Number(real) || 0);
        const p = Math.abs(Number(presupuesto) || 0);

        if (!p && r) return 999;
        if (!p && !r) return 0;

        return (r / p) * 100;
      };

      const pctSigned = function(real, presupuesto) {
        const r = Number(real) || 0;
        const p = Number(presupuesto) || 0;

        if (!p && r) return 999;
        if (!p && !r) return 0;

        return (r / p) * 100;
      };

      return {
        entradasPct: pctAbs(t.ejecutadoEntradas, t.presupuestoEntradas),
        salidasPct: pctAbs(t.ejecutadoSalidas, t.presupuestoSalidas),
        netoPct: pctSigned(t.netoEjecutado, t.netoPresupuesto),

        entradasDiff: t.ejecutadoEntradas - t.presupuestoEntradas,
        salidasDiff: t.ejecutadoSalidas - t.presupuestoSalidas,
        netoDiff: t.netoEjecutado - t.netoPresupuesto,

        noPresupuestadoSalidaPct: t.ejecutadoSalidas
          ? (t.noPresupuestadoSalida / t.ejecutadoSalidas) * 100
          : 0,

        alertasPct: this.projectionTableRowsFiltered.length
          ? (t.alertas / this.projectionTableRowsFiltered.length) * 100
          : 0
      };
    },

    // =====================================================
    // PROYECCIÓN / TAREAS PENDIENTES
    // =====================================================
    projectionPendingTasksList() {
      return Object.values(this.projectionTasks);
    },
  },

  watch: {
    'pgForm.tipo': function(nv, ov) {
      if (nv !== ov) {
        this.pgForm.categoria = '';
        this.pgForm.grupo = '';
      }
    },

    'pgForm.categoria': function(nv, ov) {
      if (nv !== ov) {
        this.pgForm.grupo = '';
      }
    },

    'pgForm.fechaRegistro': function(nv) {
      this.pgForm.fechaPago = nv;
    },

    activeView: function(nv) {
      if (this.isAuthenticated) {
        this.resetInactivityTimer();
      }

      if (nv === 'dashboard') {
        this.$nextTick(() => { this.renderCharts(); });
      }
    },

    dbLista: {
      handler: function() {
        if (this.activeView === 'dashboard') {
          this.$nextTick(() => { this.renderCharts(); });
        }
      },
      deep: true
    },

    fuentesSeleccionablesTesoreria: {
      handler: function(nuevasFuentes) {
        this.syncTesoreriaSelectedSources(nuevasFuentes);
      },
      immediate: true
    },
  },

  mounted() {
    this.checkScreenSize();
    window.addEventListener('resize', this.checkScreenSize);
    this.verifyInitialUser();
    this.fetchTRM();
  },

  beforeDestroy() {
    window.removeEventListener('resize', this.checkScreenSize);
    this.clearInactivityTimer();
    this.unbindActivityListeners();
  },

  methods: {
    async fetchTRM() {
      this.isFetchingTRM = true;
      try {
        const res = await fetch('https://open.er-api.com/v6/latest/USD');
        const data = await res.json();
        if (data && data.rates && data.rates.COP) {
          this.trmUSD = Math.round(data.rates.COP);
        }
      } catch (error) {
        console.error('Error al obtener la TRM:', error);
      } finally {
        this.isFetchingTRM = false;
      }
    },

    getTipoColor(tipo) {
      if (tipo === 'Ingreso') return 'text-tema-4';
      if (tipo === 'Gasto') return 'text-red-500';
      if (tipo === 'Activo') return 'text-tema-5';
      if (tipo === 'Pasivo') return 'text-orange-500';
      return 'text-tema-1';
    },

    getTipoBadgeColor(tipo) {
      if (tipo === 'Ingreso') return 'bg-green-100 text-green-700';
      if (tipo === 'Gasto') return 'bg-red-100 text-red-700';
      if (tipo === 'Activo') return 'bg-blue-100 text-blue-700';
      if (tipo === 'Pasivo') return 'bg-orange-100 text-orange-700';
      return 'bg-gray-100 text-gray-700';
    },

    getProgressBarColor(tipo) {
      if (tipo === 'Ingreso') return 'bg-green-400';
      if (tipo === 'Gasto') return 'bg-red-400';
      if (tipo === 'Activo') return 'bg-blue-400';
      if (tipo === 'Pasivo') return 'bg-orange-400';
      return 'bg-gray-400';
    },

    getValorColor(tipo) {
      if (tipo === 'Gasto') return 'text-red-500';
      if (tipo === 'Ingreso') return 'text-tema-4';
      return 'text-tema-1';
    },

    getVarColor(tipo, diff) {
      if (diff === 0) return 'text-gray-400';
      if (tipo === 'ingreso' || tipo === 'activo') return diff > 0 ? 'text-green-500' : 'text-red-500';
      if (tipo === 'gasto' || tipo === 'pasivo') return diff > 0 ? 'text-red-500' : 'text-green-500';
      return 'text-gray-400';
    },

    getVarIcon(diff) {
      return diff > 0 ? 'trending_up' : (diff < 0 ? 'trending_down' : 'trending_flat');
    },

    formatAbs(val) {
      if (!val && val !== 0) return '';
      const sign = val > 0 ? '+ ' : (val < 0 ? '- ' : '');
      return sign + '$' + this.formatInput(Math.abs(val));
    },

    formatPct(val) {
      if (!val && val !== 0) return '0.0%';
      const sign = val > 0 ? '+' : '';
      return sign + val.toFixed(1) + '%';
    },

    formatInput(val) {
      if (val === 0 || val === '0') return '0';
      if (!val) return '';
      if (typeof val === 'number') {
        return Math.round(val).toLocaleString('de-DE');
      }
      let clean = String(val).replace(/\./g, '').trim();
      if (isNaN(clean) || clean === '') return val;
      return Number(clean).toLocaleString('de-DE');
    },

    cleanInput(val) {
      if (!val) return '';
      return String(val).replace(/\./g, '');
    },

    parseToNumber(val) {
      if (!val && val !== 0) return 0;
      return Number(String(val).replace(/\./g, ''));
    },

    checkScreenSize() {
      this.isDesktopView = window.innerWidth >= 768;
      if (this.isDesktopView) this.showSidebar = false;
    },

    resetFilters() {
      this.searchQuery = '';
      this.filters = {
        tipo: '',
        grupo: '',
        categoria: '',
        tercero: '',
        medio: '',
        periodo: 'mes',
        useFechaPago: false
      };
    },

    openTesoreria() {
      this.resetFilters();
      this.activeView = 'tesoreria';
      this.fetchListData();
    },

    syncTesoreriaSelectedSources(fuentes = this.fuentesSeleccionablesTesoreria) {
      const disponibles = fuentes.map(f => f.nombre);

      if (!this.tesoreriaSelectedSources.length) {
        this.tesoreriaSelectedSources = this.tesoreriaSelectedDefaultSources.filter(nombre =>
          disponibles.includes(nombre)
        );
        return;
      }

      this.tesoreriaSelectedSources = this.tesoreriaSelectedSources.filter(nombre =>
        disponibles.includes(nombre)
      );
    },

    openDashboard() {
      this.resetFilters();
      this.activeView = 'dashboard';
      this.fetchListData();
    },

    getDiff(medio) {
      let real = this.saldosReales[medio.nombre];
      if (real === undefined || real === '') return 0;
      return this.parseToNumber(real) - medio.saldo;
    },

    async ejecutarCuadre(medio, event) {
      event.target.checked = false;

      const diff = this.getDiff(medio);
      if (diff === 0) return;

      this.$set(this.isSavingAjuste, medio.nombre, true);

      try {
        const today = new Date().toLocaleDateString('en-CA', { timeZone: 'America/Bogota' });
        const fp = this.dbFormasPago.find(f => f.medio === medio.nombre) || { clase: '', tipo_fuente: '' };

        const tipo = 'Gasto';
        const val = diff > 0 ? -Math.abs(diff) : Math.abs(diff);

        const record = {
          fecha_registro: today,
          fecha_pago: today,
          tipo: tipo,
          grupo: 'Varios',
          categoria: 'Ajuste',
          medio: medio.nombre,
          clase: fp.clase,
          tipo_fuente: fp.tipo_fuente,
          tercero: 'Luis',
          valor: val,
          descripcion: 'Ajuste',
          estado: 'h'
        };

        const response = await apiPost('saveTransaction', { record });

        if (!response.success) {
          this.failedErrorMessage = response.message || 'No se pudo guardar el ajuste';
          this.showErrorModal = true;
          return;
        }

        this.$set(this.saldosReales, medio.nombre, '');
        await this.fetchListData(true);
      } catch (error) {
        this.failedErrorMessage = 'Error de conexión guardando ajuste';
        this.showErrorModal = true;
        console.error(error);
      } finally {
        this.$set(this.isSavingAjuste, medio.nombre, false);
      }
    },

    renderCharts() {
      const ctxPyg = document.getElementById('chartPyg');
      const ctxPasivos = document.getElementById('chartPasivos');

      if (!ctxPyg || !ctxPasivos) return;

      if (this.chartPygInstance) this.chartPygInstance.destroy();
      if (this.chartPasivosInstance) this.chartPasivosInstance.destroy();

      const cData = this.dashboardChartsData;

      const compactFormatter = function(value) {
        return new Intl.NumberFormat('es-CO', {
          notation: 'compact',
          compactDisplay: 'short'
        }).format(value);
      };

      this.chartPygInstance = new Chart(ctxPyg, {
        type: 'bar',
        data: {
          labels: cData.labels,
          datasets: [
            {
              label: '% Utilidad',
              type: 'line',
              data: cData.utilidadPct,
              borderColor: '#56AED9',
              backgroundColor: '#56AED9',
              borderWidth: 2,
              yAxisID: 'y1',
              tension: 0.3,
              pointRadius: 3
            },
            {
              label: 'Ingresos',
              data: cData.ingresos,
              backgroundColor: '#87D8C2',
              borderRadius: 2,
              yAxisID: 'y'
            },
            {
              label: 'Gastos',
              data: cData.gastos,
              backgroundColor: '#EF4444',
              borderRadius: 2,
              yAxisID: 'y'
            }
          ]
        },
        options: {
          responsive: true,
          maintainAspectRatio: false,
          interaction: { mode: 'index', intersect: false },
          plugins: {
            legend: { position: 'top', labels: { boxWidth: 10, font: { size: 10 } } },
            tooltip: {
              callbacks: {
                label: function(context) {
                  let label = context.dataset.label || '';
                  if (label) label += ': ';
                  if (context.dataset.yAxisID === 'y1') label += context.parsed.y.toFixed(1) + '%';
                  else label += '$' + new Intl.NumberFormat('es-CO').format(context.parsed.y);
                  return label;
                }
              }
            }
          },
          scales: {
            x: { grid: { display: false }, ticks: { font: { size: 9 } } },
            y: {
              type: 'linear',
              display: true,
              position: 'right',
              grid: { borderDash: [2, 4], color: '#E5E7EB' },
              ticks: { font: { size: 9 }, callback: compactFormatter }
            },
            y1: {
              type: 'linear',
              display: true,
              position: 'left',
              grid: { display: false },
              ticks: { font: { size: 9 }, callback: function(v) { return v + '%'; } }
            }
          }
        }
      });

      this.chartPasivosInstance = new Chart(ctxPasivos, {
        type: 'bar',
        data: {
          labels: cData.labels,
          datasets: [
            {
              label: 'Saldo Acumulado',
              type: 'line',
              data: cData.pasivosSaldo,
              borderColor: '#F97316',
              backgroundColor: 'rgba(249, 115, 22, 0.1)',
              borderWidth: 2,
              fill: true,
              tension: 0.3,
              pointRadius: 2,
              yAxisID: 'y'
            },
            {
              label: 'Movimiento del Mes',
              data: cData.pasivosMov,
              backgroundColor: '#D1D5DB',
              borderRadius: 2,
              yAxisID: 'y1'
            }
          ]
        },
        options: {
          responsive: true,
          maintainAspectRatio: false,
          interaction: { mode: 'index', intersect: false },
          plugins: {
            legend: { position: 'top', labels: { boxWidth: 10, font: { size: 10 } } },
            tooltip: {
              callbacks: {
                label: function(context) {
                  let label = context.dataset.label || '';
                  if (label) label += ': ';
                  label += '$' + new Intl.NumberFormat('es-CO').format(context.parsed.y);
                  return label;
                }
              }
            }
          },
          scales: {
            x: { grid: { display: false }, ticks: { font: { size: 9 } } },
            y: {
              type: 'linear',
              display: true,
              position: 'left',
              grid: { borderDash: [2, 4], color: '#E5E7EB' },
              ticks: { font: { size: 9 }, callback: compactFormatter }
            },
            y1: {
              type: 'linear',
              display: true,
              position: 'right',
              grid: { display: false },
              ticks: { font: { size: 9 }, callback: compactFormatter }
            }
          }
        }
      });
    },

    async verifyInitialUser() {
      if (this.isAuthenticated && (this.activeView === 'dashboard' || this.activeView === 'tesoreria' || this.activeView === 'lista')) {
        this.isLoadingList = true;
      }

      this.isCheckingUser = false;

      if (this.isAuthenticated) {
        await this.fetchFormData();

        await this.fetchFormData();
          this.logAccess();

          if (this.activeView === 'dashboard') {
            await this.fetchInitialDashboardData();
          } else if (this.activeView === 'tesoreria' || this.activeView === 'lista') {
            await this.fetchListData();
          }

        this.bindActivityListeners();
        this.startInactivityTimer();
      }
    },

    async logAccess() {
      try {
        await apiPost('logAccess', {
          username: this.loginForm.username || this.userName || 'usuario'
        });
      } catch (error) {
        console.error('No se pudo registrar acceso:', error);
      }
    },

    async handleLogin() {
      this.isLoggingIn = true;
      this.loginError = '';

      try {
        const response = await apiPost('validateLogin', {
          username: this.loginForm.username,
          password: this.loginForm.password
        });

        this.isLoggingIn = false;

        if (response.success) {
          this.userName = response.data.user.name || '';

          if (this.activeView === 'dashboard' || this.activeView === 'tesoreria' || this.activeView === 'lista') {
            this.isLoadingList = true;
          }

          if (this.activeView === 'proyeccion') {
            this.isLoadingProjection = true;
          }

          this.isAuthenticated = true;

          if (this.activeView === 'dashboard') {
            this.fetchFormData();
            this.logAccess();
            await this.fetchInitialDashboardData();
          } else {
            await this.fetchFormData();
            this.logAccess();

            if (this.activeView === 'tesoreria' || this.activeView === 'lista') {
              await this.fetchListData();
            }

            if (this.activeView === 'proyeccion') {
              await this.fetchProjectionData();
            }
          }

          this.bindActivityListeners();
          this.startInactivityTimer();
        } else {
          this.loginError = response.message || 'Credenciales inválidas';
        }
      } catch (error) {
        this.isLoggingIn = false;
        this.loginError = 'No se pudo conectar con el servidor';
        console.error(error);
      }
    },

    logout() {
      this.clearInactivityTimer();
      this.unbindActivityListeners();

      this.isAuthenticated = false;
      this.userName = '';
      this.userEmail = '';
      this.loginForm.password = '';
    },

    async fetchFormData() {
      this.isLoadingFormData = true;

      try {
        const response = await apiGet('getFormData');

        if (response.success) {
          this.dbTerceros = response.data.terceros || [];
          this.dbFormasPago = response.data.formasPago || [];
          this.dbTablaPG = response.data.tablaPG || [];
          this.dbFlash = response.data.flashData || [];
        } else {
          this.failedErrorMessage = response.message || 'No se pudo cargar la información del formulario';
          this.showErrorModal = true;
        }
      } catch (error) {
        this.failedErrorMessage = 'Error de conexión cargando datos del formulario';
        this.showErrorModal = true;
        console.error(error);
      } finally {
        this.isLoadingFormData = false;
      }
    },

    async savePG() {
      this.isSavingPG = true;

      try {
        const fp = this.dbFormasPago.find(f => f.medio === this.pgForm.formaPago) || { clase: '', tipo_fuente: '' };
        const valorLimpio = this.parseToNumber(this.pgForm.valor);

        const record = {
          fecha_registro: this.pgForm.fechaRegistro,
          fecha_pago: this.pgForm.fechaPago,
          tipo: this.pgForm.tipo,
          grupo: this.pgForm.grupo,
          categoria: this.pgForm.categoria,
          medio: this.pgForm.formaPago,
          clase: fp.clase,
          tipo_fuente: fp.tipo_fuente,

          genera_pasivo: fp.genera_pasivo || 'n',
          tercero_fuente: fp.tercero_fuente || '',

          tercero: this.pgForm.tercero,
          valor: valorLimpio,
          descripcion: this.pgForm.descripcion,
          estado: 'h'
          
        };

        const response = await apiPost('saveTransaction', { record });

        if (!response.success) {
          this.failedErrorMessage = response.message || 'No se pudo guardar';
          this.showErrorModal = true;
          return;
        }

        this.isSaveSuccess = true;
        this.pgForm.valor = '';
        this.pgForm.descripcion = '';

        setTimeout(() => {
          this.isSaveSuccess = false;
        }, 2000);

      } catch (error) {
        this.failedErrorMessage = 'Error de conexión guardando la transacción';
        this.showErrorModal = true;
        console.error(error);
      } finally {
        this.isSavingPG = false;
      }
    },

    async saveFlashPG(item) {
      item.isSaving = true;

      try {
        const today = new Date().toLocaleDateString('en-CA', { timeZone: 'America/Bogota' });
        const fp = this.dbFormasPago.find(f => f.medio === item.medio) || { clase: '', tipo_fuente: '' };
        const valorLimpio = this.parseToNumber(item.valor);

        const record = {
          fecha_registro: today,
          fecha_pago: today,
          tipo: item.tipo,
          grupo: item.grupo,
          categoria: item.categoria,
          medio: item.medio,
          clase: fp.clase,
          tipo_fuente: fp.tipo_fuente,

          genera_pasivo: fp.genera_pasivo || 'n',
          tercero_fuente: fp.tercero_fuente || '',

          tercero: item.tercero,
          valor: valorLimpio,
          descripcion: item.descripcion,
          estado: 'h'
          
        };

        const response = await apiPost('saveTransaction', { record });

        if (!response.success) {
          this.failedErrorMessage = response.message || 'No se pudo guardar el registro flash';
          this.showErrorModal = true;
          return;
        }

        item.valor = '';
        item.checked = false;
      } catch (error) {
        this.failedErrorMessage = 'Error de conexión guardando el registro flash';
        this.showErrorModal = true;
        console.error(error);
      } finally {
        item.isSaving = false;
      }
    },

    async fetchListData(silent = false) {
      if (!silent) {
        this.isLoadingList = true;
        this.dbLista = [];
      } else {
        this.isSyncingBackground = true;
      }

      try {
        const response1 = await apiGet('getListData', { onlyCurrentYear: true });

        if (response1.success) {
          this.dbLista = response1.data.items || [];
        }

        if (!silent) {
          this.isLoadingList = false;
        }

        this.isSyncingBackground = true;

        const response2 = await apiGet('getListData', { onlyCurrentYear: false });

        if (response2.success) {
          this.dbLista = response2.data.items || [];
        }
      } catch (error) {
        console.error(error);
        this.failedErrorMessage = 'Error cargando movimientos';
        this.showErrorModal = true;

        if (!silent) {
          this.isLoadingList = false;
        }
      } finally {
        this.isSyncingBackground = false;
        this.isLoadingList = false;
      }
    },

    async deleteTransaction(transaction) {
      const transactionId = transaction.idx || transaction.id;
      const index = this.dbLista.findIndex(t => (t.idx || t.id) == transactionId);

      if (index === -1) return;

      const backup = this.dbLista.splice(index, 1)[0];

      try {
        const response = await apiPost('softDeleteTransaction', {
          id: transactionId
        });

        if (!response.success) {
          this.dbLista.splice(index, 0, backup);
          this.failedErrorMessage = response.message || 'No se pudo eliminar';
          this.showErrorModal = true;
        }
      } catch (error) {
        this.dbLista.splice(index, 0, backup);
        this.failedErrorMessage = 'Error de conexión eliminando transacción';
        this.showErrorModal = true;
        console.error(error);
      }
    },

    startInactivityTimer() {
      this.clearInactivityTimer();

      this.inactivityTimer = setTimeout(() => {
        this.forceAutoLogout();
      }, this.inactivityTimeoutMs);
    },

    resetInactivityTimer() {
      if (!this.isAuthenticated) return;
      this.startInactivityTimer();
    },

    clearInactivityTimer() {
      if (this.inactivityTimer) {
        clearTimeout(this.inactivityTimer);
        this.inactivityTimer = null;
      }
    },

    handleUserActivity() {
      if (!this.isAuthenticated) return;
      this.resetInactivityTimer();
    },

    bindActivityListeners() {
      if (this.activityListenersBound) return;

      const events = ['click', 'keydown', 'input', 'change', 'touchstart', 'scroll'];

      events.forEach(eventName => {
        window.addEventListener(eventName, this.handleUserActivity, true);
      });

      this.activityListenersBound = true;
    },

    unbindActivityListeners() {
      if (!this.activityListenersBound) return;

      const events = ['click', 'keydown', 'input', 'change', 'touchstart', 'scroll'];

      events.forEach(eventName => {
        window.removeEventListener(eventName, this.handleUserActivity, true);
      });

      this.activityListenersBound = false;
    },

    forceAutoLogout() {
      this.logout();
    },

    toggleFuenteTesoreria(nombre) {
      const index = this.tesoreriaSelectedSources.indexOf(nombre);

      if (index === -1) {
        this.tesoreriaSelectedSources.push(nombre);
      } else {
        this.tesoreriaSelectedSources.splice(index, 1);
      }
    },

    async fetchInitialDashboardData() {
      this.isLoadingList = true;
      this.isSyncingBackground = false;
      this.dbLista = [];

      try {
        const response1 = await apiGet('getListData', { onlyCurrentYear: true });

        if (response1.success) {
          this.dbLista = response1.data.items || [];
        }

        this.isLoadingList = false;

        this.$nextTick(() => {
          if (this.activeView === 'dashboard') {
            this.renderCharts();
          }
        });

        this.isSyncingBackground = true;

        apiGet('getListData', { onlyCurrentYear: false })
          .then(response2 => {
            if (response2.success) {
              this.dbLista = response2.data.items || [];
            }

            this.$nextTick(() => {
              if (this.activeView === 'dashboard') {
                this.renderCharts();
              }
            });
          })
          .catch(error => {
            console.error(error);
          })
          .finally(() => {
            this.isSyncingBackground = false;
          });

      } catch (error) {
        console.error(error);
        this.failedErrorMessage = 'Error cargando datos iniciales del dashboard';
        this.showErrorModal = true;
        this.isLoadingList = false;
        this.isSyncingBackground = false;
      }
    },

    parseExpression(val) {
      if (!val) return val;

      let exp = String(val)
        .replace(/\./g, '')
        .replace(/\s+/g, '')
        .replace(/,/g, '');

      // validar caracteres permitidos
      if (!/^[0-9+\-*/]+$/.test(exp)) return val;

      try {
        // 🔹 Manejar negativos al inicio
        if (exp[0] === '-') exp = '0' + exp;

        // 🔹 Tokenizar
        let tokens = exp.match(/(\d+|\+|\-|\*|\/)/g);

        if (!tokens) return val;

        // 🔹 Paso 1: resolver * y /
        let stack = [];
        let i = 0;

        while (i < tokens.length) {
          let token = tokens[i];

          if (token === '*' || token === '/') {
            let prev = Number(stack.pop());
            let next = Number(tokens[i + 1]);

            let res = token === '*'
              ? prev * next
              : (next === 0 ? 0 : prev / next);

            stack.push(res);
            i += 2;
          } else {
            stack.push(token);
            i++;
          }
        }

        // 🔹 Paso 2: resolver + y -
        let result = Number(stack[0]);

        for (let j = 1; j < stack.length; j += 2) {
          let operator = stack[j];
          let num = Number(stack[j + 1]);

          if (operator === '+') result += num;
          else result -= num;
        }

        return Math.round(result);

      } catch (e) {
        return val;
      }
    },

    handleValorBlur(val) {
      if (!val) return '';

      // 1. intentar evaluar
      const calculado = this.parseExpression(val);

      // 2. formatear resultado
      return this.formatInput(calculado);
    },

    // =====================================================
    // PROYECCIÓN / APERTURA DEL MÓDULO
    // =====================================================
    async openProyeccion() {
      this.activeView = 'proyeccion';

      if (!this.isDesktopView) {
        this.showSidebar = false;
      }

      if (!this.dbTablaPG.length || !this.dbTerceros.length) {
        this.fetchFormData();
      }

      await this.fetchProjectionData();
    },

    // =====================================================
    // PROYECCIÓN / CARGA REAL DE DATOS
    // Carga primero año actual y luego histórico en background
    // =====================================================
    async fetchProjectionData(silent = false) {
      if (!silent) {
        this.isLoadingProjection = true;
        this.dbProjectionBudget = [];
        this.dbProjectionExecution = [];
      } else {
        this.isSyncingProjectionBackground = true;
      }

      try {
        const response1 = await apiGet('getProjectionData', {
          onlyCurrentYear: true
        });

        console.group('PROYECCIÓN / FETCH AÑO ACTUAL');
        console.log('Response current year:', response1);
        console.table(response1.data && response1.data.budget ? response1.data.budget : []);
        console.groupEnd();

        if (!response1.success) {
          this.failedErrorMessage = response1.message || 'No se pudo cargar la proyección';
          this.showErrorModal = true;
          return;
        }

        this.dbProjectionBudget = response1.data.budget || [];
        this.dbProjectionExecution = response1.data.execution || [];

        if (!silent) {
          this.isLoadingProjection = false;
        }

        this.isSyncingProjectionBackground = true;

        const response2 = await apiGet('getProjectionData', {
          onlyCurrentYear: false
        });

        console.group('PROYECCIÓN / FETCH HISTÓRICO');
        console.log('Response histórico:', response2);
        console.table(response2.data && response2.data.budget ? response2.data.budget : []);
        console.groupEnd();

        if (response2.success) {
          this.dbProjectionBudget = response2.data.budget || [];
          this.dbProjectionExecution = response2.data.execution || [];
        } else {
          console.warn('No se pudo cargar histórico de proyección:', response2.message);
        }

      } catch (error) {
        console.error(error);
        this.failedErrorMessage = 'Error de conexión cargando datos de proyección';
        this.showErrorModal = true;
      } finally {
        this.isLoadingProjection = false;
        this.isSyncingProjectionBackground = false;
      }
    },

    // =====================================================
    // PROYECCIÓN / FECHAS Y KEYS
    // =====================================================
    getProjectionDateRange(periodo) {
      const today = new Date();
      const y = today.getFullYear();
      const m = today.getMonth();
      const d = today.getDate();

      let start = null;
      let end = null;

      switch (periodo) {
        case 'mes':
          start = new Date(y, m, 1);
          end = new Date(y, m + 1, 0);
          break;
        
        case 'mes_siguiente':
          start = new Date(y, m + 1, 1);
          end = new Date(y, m + 2, 0);
          break;

        case 'mes_anterior':
          start = new Date(y, m - 1, 1);
          end = new Date(y, m, 0);
          break;

        case 'ytd':
          start = new Date(y, 0, 1);
          end = new Date(y, m, d);
          break;

        case 'año':
          start = new Date(y, 0, 1);
          end = new Date(y, 11, 31);
          break;

        case 'año_anterior':
          start = new Date(y - 1, 0, 1);
          end = new Date(y - 1, 11, 31);
          break;

        default:
          return { start: null, end: null };
      }

      start.setHours(0, 0, 0, 0);
      end.setHours(23, 59, 59, 999);

      return { start, end };
    },

    // =====================================================
    // PROYECCIÓN / FECHA DEFAULT PARA NUEVO PRESUPUESTO
    // Usa una fecha dentro del periodo visible
    // =====================================================
    getProjectionDefaultBudgetDate() {
      const range = this.getProjectionDateRange(this.projectionFilters.periodo);

      if (range && range.start) {
        return range.start.toLocaleDateString('en-CA', { timeZone: 'America/Bogota' });
      }

      return new Date().toLocaleDateString('en-CA', { timeZone: 'America/Bogota' });
    },

    isProjectionDateInRange(dateStr, range) {
      if (!range || (!range.start && !range.end)) return true;
      if (!dateStr) return false;

      const d = new Date(dateStr + 'T00:00:00');

      return d >= range.start && d <= range.end;
    },

    getProjectionMonthRangeFromDate(dateStr) {
      const d = new Date(dateStr + 'T00:00:00');
      const start = new Date(d.getFullYear(), d.getMonth(), 1);
      const end = new Date(d.getFullYear(), d.getMonth() + 1, 0);

      start.setHours(0, 0, 0, 0);
      end.setHours(23, 59, 59, 999);

      return { start, end };
    },

    normalizeProjectionValue(value) {
      return String(value || '')
        .toLowerCase()
        .normalize('NFD')
        .replace(/[\u0300-\u036f]/g, '')
        .trim();
    },

    // =====================================================
    // PROYECCIÓN / KEY DE COMPARACIÓN
    // Incluye flujo para no mezclar entradas con salidas
    // =====================================================
    buildProjectionMatchKey(tipo, grupo, categoria, tercero, flujo) {
      return [
        this.normalizeProjectionValue(tipo),
        this.normalizeProjectionValue(grupo),
        this.normalizeProjectionValue(categoria),
        this.normalizeProjectionValue(tercero),
        this.normalizeProjectionValue(flujo)
      ].join('|');
    },

    // =====================================================
    // PROYECCIÓN / SENTIDO DEL FLUJO DE CAJA
    // =====================================================
    getProjectionFlowDirection(tipo, valor) {
      const tipoNorm = this.normalizeProjectionValue(tipo);
      const num = Number(valor) || 0;

      // Ingreso siempre es entrada, incluso si viene negativo
      if (tipoNorm === 'ingreso') {
        return 'entrada';
      }

      // Gasto siempre es salida, incluso si viene negativo
      if (tipoNorm === 'gasto') {
        return 'salida';
      }

      // Activo positivo = constitución/aumento = salida
      // Activo negativo = reintegro/retiro = entrada
      if (tipoNorm === 'activo') {
        return num < 0 ? 'entrada' : 'salida';
      }

      // Pasivo positivo = entrada de caja por deuda
      // Pasivo negativo = pago/disminución de deuda = salida
      if (tipoNorm === 'pasivo') {
        return num < 0 ? 'salida' : 'entrada';
      }

      return 'salida';
    },

    getProjectionFlowLabel(flujo) {
      return flujo === 'entrada' ? 'Entrada' : 'Salida';
    },

    getProjectionFlowBadgeClass(flujo) {
      if (flujo === 'entrada') {
        return 'bg-green-100 text-green-700';
      }

      return 'bg-red-100 text-red-700';
    },

    // =====================================================
    // PROYECCIÓN / VALOR COMPARABLE
    // Para comparar presupuesto vs ejecución usamos magnitud.
    // El sentido real lo define getProjectionFlowDirection().
    // =====================================================
    getProjectionComparableValue(tipo, valor) {
      const num = Number(valor) || 0;
      return Math.abs(num);
    },


    // =====================================================
    // PROYECCIÓN / ALERTAS POR FLUJO
    // Entrada: más ejecución es bueno.
    // Salida: pasarse del presupuesto es malo.
    // =====================================================
    getProjectionAlert(flujo, porcentaje, presupuesto, ejecutado) {
      if (!presupuesto && ejecutado > 0) {
        if (flujo === 'entrada') {
          return {
            level: 'ok',
            label: 'Entrada extra',
            className: 'bg-green-100 text-green-700',
            barClass: 'bg-green-400'
          };
        }

        return {
          level: 'danger',
          label: 'Sin ppto',
          className: 'bg-red-600 text-white shadow-sm',
          barClass: 'bg-red-600'
        };
      }

      if (!presupuesto && !ejecutado) {
        return {
          level: 'neutral',
          label: '0%',
          className: 'bg-gray-100 text-gray-500',
          barClass: 'bg-gray-300'
        };
      }

      // Entradas: alcanzar o superar presupuesto es positivo
      if (flujo === 'entrada') {
        if (porcentaje >= 100) {
          return {
            level: 'ok',
            label: 'OK',
            className: 'bg-green-100 text-green-700',
            barClass: 'bg-green-400'
          };
        }

        if (porcentaje >= 90) {
          return {
            level: 'warning',
            label: 'Cerca',
            className: 'bg-orange-100 text-orange-700',
            barClass: 'bg-orange-400'
          };
        }

        return {
          level: 'danger',
          label: 'Bajo',
          className: 'bg-red-100 text-red-700',
          barClass: 'bg-red-400'
        };
      }

      // Salidas: consumir menos presupuesto es positivo
      if (porcentaje < 70) {
        return {
          level: 'ok',
          label: 'Verde',
          className: 'bg-green-100 text-green-700',
          barClass: 'bg-green-400'
        };
      }

      if (porcentaje < 95) {
        return {
          level: 'warning',
          label: 'Naranja',
          className: 'bg-orange-100 text-orange-700',
          barClass: 'bg-orange-400'
        };
      }

      if (porcentaje <= 110) {
        return {
          level: 'danger',
          label: 'Rojo',
          className: 'bg-red-100 text-red-700',
          barClass: 'bg-red-400'
        };
      }

      return {
        level: 'danger',
        label: 'Crítico',
        className: 'bg-red-600 text-white shadow-sm',
        barClass: 'bg-red-600'
      };
    },

    // =====================================================
    // PROYECCIÓN / COLOR DIFERENCIA POR FLUJO
    // Entrada: ejecutar más es bueno.
    // Salida: ejecutar más es malo.
    // =====================================================
    getProjectionDiffColor(row) {
      if (row.diferencia === 0) return 'text-gray-400';

      if (row.flujo === 'entrada') {
        return row.diferencia > 0 ? 'text-tema-4' : 'text-red-500';
      }

      return row.diferencia > 0 ? 'text-red-500' : 'text-tema-4';
    },

    getProjectionDetailTooltip(row) {
      if (!row.detalle || !row.detalle.length) return '';

      return row.detalle.map(item => {
        return [
          item.categoria || 'Sin categoría',
          item.tercero || 'Sin tercero',
          '$ ' + this.formatInput(item.valor)
        ].join(' • ');
      }).join('\n');
    },

    // =====================================================
    // PROYECCIÓN / OPCIONES POR FILA EDITABLE
    // =====================================================
    projectionGrupoOptionsByTipo(tipo) {
      return [...new Set(
        this.dbTablaPG
          .filter(item => !tipo || item.tipo === tipo)
          .map(item => item.grupo)
      )].filter(Boolean).sort();
    },

    projectionCategoriaOptionsByTipoGrupo(tipo, grupo) {
      return [...new Set(
        this.dbTablaPG
          .filter(item => !tipo || item.tipo === tipo)
          .filter(item => !grupo || item.grupo === grupo)
          .map(item => item.categoria)
      )].filter(Boolean).sort();
    },

    // =====================================================
    // PROYECCIÓN / EDICIÓN DE FILAS
    // =====================================================
    isProjectionRowEditing(row) {
      return !!this.projectionEditingRows[row.key];
    },

    // =====================================================
    // PROYECCIÓN / NUEVA FILA INDIVIDUAL
    // Crea una fila pendiente y la deja editable en línea
    // =====================================================
    addProjectionSingleRow() {
      const rowKey = 'new-' + Date.now();

      const record = {
        fecha_ejecucion: this.getProjectionDefaultBudgetDate(),

        tipo: this.projectionFilters.tipo || '',
        grupo: this.projectionFilters.grupo || '',
        categoria: this.projectionFilters.categoria || '',
        tercero: this.projectionFilters.tercero || '',

        valor: 0,
        descripcion: '',
        estado: 'h'
      };

      this.$set(this.projectionTasks, rowKey, {
        action: 'create',
        id: null,
        rowKey: rowKey,
        record: record
      });

      this.$set(this.projectionDrafts, rowKey, {
        fecha_ejecucion: record.fecha_ejecucion,
        tipo: record.tipo,
        grupo: record.grupo,
        categoria: record.categoria,
        tercero: record.tercero,
        valor: this.formatInput(record.valor),
        descripcion: record.descripcion
      });

      this.$set(this.projectionEditingRows, rowKey, true);
    },

    editProjectionRow(row) {
      const today = new Date().toLocaleDateString('en-CA', { timeZone: 'America/Bogota' });

      const defaultCategoria = row.isVirtual && row.categoria === 'No presupuestado'
        ? ''
        : (row.categoria || '');

      const defaultTercero = row.isVirtual && row.tercero === 'No presupuestado'
        ? ''
        : (row.tercero || '');

      this.$set(this.projectionDrafts, row.key, {
        fecha_ejecucion: row.fecha_ejecucion || today,
        tipo: row.tipo || '',
        grupo: row.grupo || '',
        categoria: defaultCategoria,
        tercero: defaultTercero,
        valor: this.formatInput(row.presupuesto || row.ejecutado || 0),
        descripcion: row.descripcion || ''
      });

      this.$set(this.projectionEditingRows, row.key, true);

      if (row.isVirtual) {
        this.queueProjectionDraft(row);
      }
    },

    closeProjectionRowEdit(row) {
      this.$delete(this.projectionEditingRows, row.key);
    },

    // =====================================================
    // PROYECCIÓN / CANCELAR EDICIÓN O CREACIÓN
    // Si es nueva o viene desde +, elimina la tarea local.
    // Si es existente, revierte cambios pendientes.
    // =====================================================
    cancelProjectionRowEdit(row) {
      this.$delete(this.projectionDrafts, row.key);
      this.$delete(this.projectionEditingRows, row.key);
      this.$delete(this.projectionDeletedRows, row.key);

      const task = this.projectionTasks[row.key];

      if (task) {
        this.$delete(this.projectionTasks, row.key);
      }
    },

    onProjectionDraftTipoChange(row) {
      if (!this.projectionDrafts[row.key]) return;

      this.projectionDrafts[row.key].grupo = '';
      this.projectionDrafts[row.key].categoria = '';

      this.queueProjectionDraft(row);
    },

    onProjectionDraftGrupoChange(row) {
      if (!this.projectionDrafts[row.key]) return;

      this.projectionDrafts[row.key].categoria = '';

      this.queueProjectionDraft(row);
    },

    queueProjectionDraft(row) {
      const draft = this.projectionDrafts[row.key];

      if (!draft) return;

      const record = {
        fecha_ejecucion: draft.fecha_ejecucion,
        tipo: draft.tipo,
        grupo: draft.grupo,
        categoria: draft.categoria,
        tercero: draft.tercero,
        valor: this.parseToNumber(draft.valor),
        descripcion: draft.descripcion,
        estado: 'h'
      };

      const action = row.isVirtual || row.isPendingCreate || !row.budgetId
        ? 'create'
        : 'update';

      this.$set(this.projectionTasks, row.key, {
        action: action,
        id: action === 'update' ? row.budgetId : null,
        rowKey: row.key,
        record: record
      });
    },

    deleteProjectionBudgetRow(row) {
      if (row.isVirtual) return;

      // Si es una fila nueva pendiente, no existe en back.
      // Entonces simplemente se cancela localmente.
      if (row.isPendingCreate || !row.budgetId) {
        this.$delete(this.projectionTasks, row.key);
        this.$delete(this.projectionDrafts, row.key);
        this.$delete(this.projectionEditingRows, row.key);
        this.$delete(this.projectionDeletedRows, row.key);
        return;
      }

      this.$set(this.projectionDeletedRows, row.key, true);

      this.$set(this.projectionTasks, row.key, {
        action: 'delete',
        id: row.budgetId,
        rowKey: row.key
      });
    },

    restoreProjectionBudgetRow(row) {
      this.$delete(this.projectionDeletedRows, row.key);
      this.$delete(this.projectionTasks, row.key);
    },

    enableProjectionBulkEdit() {
      this.projectionTableRowsFiltered.forEach(row => {
        if (!this.projectionDeletedRows[row.key]) {
          this.editProjectionRow(row);
        }
      });
    },


    // =====================================================
    // PROYECCIÓN / VALIDACIÓN DE TAREAS ANTES DE BACK
    // Obligatorios: fecha_ejecucion, tipo, grupo, valor, descripcion
    // Opcionales: categoria, tercero
    // =====================================================
    validateProjectionTasksBeforeSave() {
      const errors = [];

      this.projectionPendingTasksList.forEach((task, index) => {
        if (task.action === 'delete') return;

        const record = task.record || {};

        if (!record.fecha_ejecucion) {
          errors.push(`Tarea ${index + 1}: fecha_ejecucion es requerida`);
        }

        if (!record.tipo) {
          errors.push(`Tarea ${index + 1}: tipo es requerido`);
        }

        if (!record.grupo) {
          errors.push(`Tarea ${index + 1}: grupo es requerido`);
        }

        if (
          record.valor === '' ||
          record.valor === null ||
          record.valor === undefined ||
          isNaN(Number(record.valor))
        ) {
          errors.push(`Tarea ${index + 1}: valor debe ser numérico`);
        }

        if (!record.descripcion || !String(record.descripcion).trim()) {
          errors.push(`Tarea ${index + 1}: descripción es requerida`);
        }
      });

      return errors;
    },

    // =====================================================
    // PROYECCIÓN / GUARDADO REAL EN BACK
    // Envía creates, updates y deletes al Apps Script
    // =====================================================
    async saveProjectionTasks() {
      if (!this.projectionPendingTasksList.length) return;

      const validationErrors = this.validateProjectionTasksBeforeSave();

      if (validationErrors.length > 0) {
        this.failedErrorMessage = validationErrors.slice(0, 8).join('\n');
        this.showErrorModal = true;
        return;
      }

      this.isSavingProjection = true;
      this.projectionSaveSuccess = false;

      try {
        console.group('PROYECCIÓN / GUARDAR');
        console.log('Tasks enviadas:', JSON.parse(JSON.stringify(this.projectionPendingTasksList)));

        const response = await apiPost('saveProjectionTasks', {
          tasks: this.projectionPendingTasksList
        });

        console.log('Respuesta saveProjectionTasks:', response);
        console.groupEnd();

        if (!response.success) {
          this.failedErrorMessage = response.message || 'No se pudieron guardar los cambios de presupuesto';
          this.showErrorModal = true;
          return;
        }

        if (response.data && response.data.errors && response.data.errors.length > 0) {
          this.failedErrorMessage = response.data.errors
            .slice(0, 8)
            .map(err => err.message || JSON.stringify(err))
            .join('\n');

          this.showErrorModal = true;
        }

        this.projectionTasks = {};
        this.projectionDrafts = {};
        this.projectionEditingRows = {};
        this.projectionDeletedRows = {};

        this.projectionSaveSuccess = true;

        await this.fetchProjectionData(true);

        setTimeout(() => {
          this.projectionSaveSuccess = false;
        }, 1800);

      } catch (error) {
        console.error(error);
        this.failedErrorMessage = 'Error de conexión guardando presupuesto';
        this.showErrorModal = true;
      } finally {
        this.isSavingProjection = false;
      }
    },

    // =====================================================
    // PROYECCIÓN / COPIAR PRESUPUESTO
    // =====================================================
    openProjectionCopyModal() {
      this.projectionCopyModal.show = true;
      this.projectionCopyPreview = [];
      this.projectionCopyModal.targetDate = new Date().toLocaleDateString('en-CA', { timeZone: 'America/Bogota' });
    },

    closeProjectionCopyModal() {
      this.projectionCopyModal.show = false;
      this.projectionCopyPreview = [];
    },

    prepareProjectionCopyPreview() {
      const targetDate = this.projectionCopyModal.targetDate;

      if (!targetDate) {
        this.failedErrorMessage = 'Selecciona una fecha destino para copiar el presupuesto';
        this.showErrorModal = true;
        return;
      }

      const targetRange = this.getProjectionMonthRangeFromDate(targetDate);

      const existsInTargetMonth = this.dbProjectionBudget.some(item => {
        if ((item.estado || 'h') === 'n') return false;
        return this.isProjectionDateInRange(item.fecha_ejecucion, targetRange);
      });

      if (existsInTargetMonth) {
        this.failedErrorMessage = 'Ya existe presupuesto activo en el mes destino. No se generó la copia para evitar duplicados.';
        this.showErrorModal = true;
        return;
      }

      const sourceRange = this.getProjectionDateRange(this.projectionFilters.periodo);

      const sourceRows = this.dbProjectionBudget.filter(item => {
        if ((item.estado || 'h') === 'n') return false;
        return this.isProjectionDateInRange(item.fecha_ejecucion, sourceRange);
      });

      if (!sourceRows.length) {
        this.failedErrorMessage = 'No hay registros presupuestados en el periodo visible para copiar.';
        this.showErrorModal = true;
        return;
      }

      this.projectionCopyPreview = sourceRows.map((item, index) => {
        return {
          tmpId: 'copy-' + index + '-' + Date.now(),
          fecha_ejecucion: targetDate,
          tipo: item.tipo || '',
          grupo: item.grupo || '',
          categoria: item.categoria || '',
          tercero: item.tercero || '',
          valor: this.formatInput(item.valor || 0),
          descripcion: item.descripcion || ''
        };
      });
    },

    addProjectionCopyPreviewToTasks() {
      this.projectionCopyPreview.forEach(item => {
        const rowKey = item.tmpId;

        this.$set(this.projectionTasks, rowKey, {
          action: 'create',
          id: null,
          rowKey: rowKey,
          record: {
            fecha_ejecucion: item.fecha_ejecucion,
            tipo: item.tipo,
            grupo: item.grupo,
            categoria: item.categoria,
            tercero: item.tercero,
            valor: this.parseToNumber(item.valor),
            descripcion: item.descripcion,
            estado: 'h'
          }
        });
      });

      this.closeProjectionCopyModal();
    },

    // =====================================================
    // PROYECCIÓN / FORMATOS KPI
    // =====================================================
    formatProjectionSignedMoney(value) {
      const num = Number(value) || 0;

      if (num === 0) return '$ 0';

      const sign = num < 0 ? '- ' : '';

      return sign + '$ ' + this.formatInput(Math.abs(num));
    },

    formatProjectionDiffMoney(value) {
      const num = Number(value) || 0;

      if (num === 0) return '$ 0';

      const sign = num > 0 ? '+ ' : '- ';

      return sign + '$ ' + this.formatInput(Math.abs(num));
    },

    formatProjectionPct(value) {
      const num = Number(value);

      if (!Number.isFinite(num)) return '-';
      if (num >= 999) return 'Sin ppto';

      const sign = num < 0 ? '-' : '';

      return sign + Math.abs(num).toFixed(1) + '%';
    },

    getProjectionPctWidth(value) {
      const num = Number(value) || 0;

      if (num >= 999) return 100;

      return Math.min(Math.abs(num), 100);
    },

    getProjectionPctBarClass(value, mode) {
      const num = Number(value) || 0;

      if (num >= 999) return 'bg-red-500';

      if (mode === 'entrada') {
        if (num >= 100) return 'bg-green-500';
        if (num >= 90) return 'bg-orange-400';
        return 'bg-red-500';
      }

      if (mode === 'salida') {
        if (num < 70) return 'bg-green-500';
        if (num < 95) return 'bg-orange-400';
        return 'bg-red-500';
      }

      if (mode === 'neto') {
        if (num >= 100) return 'bg-green-500';
        if (num >= 80) return 'bg-orange-400';
        return 'bg-red-500';
      }

      return 'bg-tema-5';
    },

    getProjectionSignedValueClass(value) {
      const num = Number(value) || 0;

      if (num > 0) return 'text-tema-4';
      if (num < 0) return 'text-red-500';

      return 'text-gray-500';
    },

    getProjectionKpiDiffClass(value, mode) {
      const num = Number(value) || 0;

      if (num === 0) return 'text-gray-400';

      if (mode === 'entrada') {
        return num > 0 ? 'text-tema-4' : 'text-red-500';
      }

      if (mode === 'salida') {
        return num > 0 ? 'text-red-500' : 'text-tema-4';
      }

      if (mode === 'neto') {
        return num > 0 ? 'text-tema-4' : 'text-red-500';
      }

      return num > 0 ? 'text-tema-4' : 'text-red-500';
    },
  }
});