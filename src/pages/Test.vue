<template>
  <q-layout view="hHh lpR fFf">
    <div class="bg-orange-4 text-white">
      <q-toolbar>
        <q-btn flat round dense icon="receipt" size="xl" class="q-mt-xs" />
        <q-toolbar inset>
          <q-toolbar-title><strong>Re Cut</strong> Report</q-toolbar-title>
        </q-toolbar>
        <q-space />
        <q-btn
          flat
          round
          dense
          icon="logout"
          size="xl"
          class="q-mt-xs"
          @click="logout()"
        />
      </q-toolbar>
    </div>
    <div class="my-card" style="max-width: 100%">
      <div class="q-gutter-md">
        <q-card>
          <q-card-section>
            <q-banner rounded class="bg-primary text-white">
              <div class="text-center" style="font-size: 24px">
                Re Cut Report
              </div>
            </q-banner>
            <br />
            <br />
            <div class="row justify-center">
              <q-input
                class="q-px-xl"
                input-class="text-center"
                v-model="start"
              >
                <template v-slot:append>
                  <q-icon name="schedule" class="cursor-pointer">
                    <q-popup-proxy
                      ref="qDateProxy"
                      transition-show="scale"
                      transition-hide="scale"
                    >
                      <q-date
                        v-model="start"
                        color="yellow-7"
                        mask="DD/MM/YYYY"
                      >
                        <div class="row items-center justify-end">
                          <q-btn
                            v-close-popup
                            label="Close"
                            color="primary"
                            flat
                          />
                        </div>
                      </q-date>
                    </q-popup-proxy>
                  </q-icon>
                </template>
              </q-input>
              <q-input class="q-px-xl" input-class="text-center" v-model="end">
                <template v-slot:append>
                  <q-icon name="schedule" class="cursor-pointer">
                    <q-popup-proxy
                      ref="qDateProxy"
                      transition-show="scale"
                      transition-hide="scale"
                    >
                      <q-date v-model="end" color="yellow-7" mask="DD/MM/YYYY">
                        <div class="row items-center justify-end">
                          <q-btn
                            v-close-popup
                            label="Close"
                            color="primary"
                            flat
                          />
                        </div>
                      </q-date>
                    </q-popup-proxy>
                  </q-icon>
                </template>
              </q-input>

              <q-select
                class="q-px-xl"
                v-model="org"
                :options="g"
                bg-color=""
                style="width: 200px"
                label="Chose Org"
              />

              <q-btn
                size="md"
                dense
                class="q-px-xl q-py-xs"
                color="positive"
                label="Export"
                @click="exportexcel()"
              />
            </div>
          </q-card-section>
          <!--   <div class="q-pa-xl" style="min-width: 100%">
            <q-banner inline-actions rounded class="bg-orange-4 text-white">
              <h5 class="q-mt-xs">
                Re Cut Monthly Summary Report : {{ org }} {{ year }}
              </h5>
              <template v-slot:action> </template>
            </q-banner>

            <div
              class="col-12"
              style="min-width: 100%; min-height: 600px"
              id="chartDom"
            ></div>
          </div>
 -->
          <!--  <div class="q-pa-xl" style="min-width: 100%">
            <q-banner
              inline-actions
              rounded
              class="bg-light-green-3 text-white"
            >
              <h5 class="q-mt-xs">Cause of Re cut Loss</h5>
              <template v-slot:action> </template>
            </q-banner>

            <div
              class="col-12"
              style="min-width: 100%; min-height: 600px"
              id="chartDom2"
            ></div>
          </div>
 -->
          <q-separator />
        </q-card>
        <div class="q-pa-md q-gutter-sm">
          <q-dialog
            v-model="basic"
            transition-show="rotate"
            transition-hide="rotate"
          >
            <q-card style="min-width: 30%">
              <q-card-section>
                <div class="text-h6">Confirm Logout</div>
                <q-card-actions align="right">
                  <q-btn
                    color="green-4"
                    size="lg"
                    icon="done"
                    @click="confirm_log_out()"
                  />
                  <q-btn
                    color="negative"
                    size="lg"
                    icon="cancel"
                    @click="cancel_log_out()"
                  />
                </q-card-actions>
              </q-card-section>
            </q-card>
          </q-dialog>
        </div>
      </div>
    </div>
  </q-layout>
</template>

<script>
import { ref } from "vue";
import { date } from "quasar";
import axios from "axios";
import * as Excel from "exceljs";
import { useQuasar, QSpinnerGears } from "quasar";
import { saveAs } from "file-saver";
import moment from "moment";
import * as echarts from "echarts";
export default {
  data() {
    return {
      month: "",
      monthx: "",
      start: "",
      end: "",
      master_date_start: "",
      master_date_end: "",
      rowexport: [],
      total_pcs: [],
      total_yard: [],
      monthexport: [],
      keep_cut_qty: [],
      keep_defect_1: [],
      sum_total_yard: [],
      rowexport_use: [],
      row_result: [],
      row_result2: [],
      row_result_month: [],
      row_result_month_use: [],
      recent_yard: [],
      all_result_yard: [],
      all_result_yard2: [],
      all_result_yard3: [],
      all_result_yard4: [],
      all_result_yard5: [],
      all_result_yard6: [],
      all_result_yard7: [],
      all_result_yard8: [],
      all_result_yard9: [],
      all_result_yard10: [],
      all_result_yard11: [],
      all_result_pcs: [],
      all_result_pcs2: [],
      all_result_pcs3: [],
      all_result_pcs4: [],
      all_result_pcs5: [],
      all_result_pcs6: [],
      all_result_pcs7: [],
      all_result_pcs8: [],
      all_result_pcs9: [],
      all_result_pcs10: [],
      all_result_pcs11: [],
      recent_pcs: [],
      keep_total_yard: [],
      keep_total_per: [],
      keep_cut_qty_per: [],
      row_result_item_code: [],
      row_result_data_from_item_code: [],

      basic: false,
      year: "",
      g: ["NYG1", "NYG2", "NYG3", "NYG4"],
      org: "",
      column_main: [
        {
          col_name: "A",
        },
        {
          col_name: "B",
        },
        {
          col_name: "C",
        },
        {
          col_name: "D",
        },
        {
          col_name: "E",
        },
        {
          col_name: "F",
        },
        {
          col_name: "G",
        },
        {
          col_name: "H",
        },
        {
          col_name: "I",
        },
        {
          col_name: "J",
        },
        {
          col_name: "K",
        },
        {
          col_name: "L",
        },
        {
          col_name: "M",
        },
        {
          col_name: "N",
        },
        {
          col_name: "O",
        },
        {
          col_name: "P",
        },
        {
          col_name: "Q",
        },
        {
          col_name: "R",
        },
        {
          col_name: "S",
        },
        {
          col_name: "T",
        },
        {
          col_name: "U",
        },
        {
          col_name: "V",
        },
        {
          col_name: "W",
        },
        {
          col_name: "X",
        },
        {
          col_name: "Y",
        },
        {
          col_name: "Z",
        },
        {
          col_name: "AA",
        },
        {
          col_name: "AB",
        },
        {
          col_name: "AC",
        },
        {
          col_name: "AD",
        },
        {
          col_name: "AE",
        },
        {
          col_name: "AF",
        },
        {
          col_name: "AG",
        },
        {
          col_name: "AH",
        },
        {
          col_name: "AI",
        },
        {
          col_name: "AJ",
        },
        {
          col_name: "AK",
        },
        {
          col_name: "AL",
        },
        {
          col_name: "AM",
        },
        {
          col_name: "AN",
        },
        {
          col_name: "AO",
        },
      ],
      column_ws: [
        {
          col_name: "B",
        },
        {
          col_name: "C",
        },
        {
          col_name: "D",
        },
        {
          col_name: "E",
        },
        {
          col_name: "F",
        },
        {
          col_name: "G",
        },
        {
          col_name: "H",
        },
        {
          col_name: "I",
        },
        {
          col_name: "J",
        },
        {
          col_name: "K",
        },
        {
          col_name: "L",
        },
        {
          col_name: "M",
        },
        {
          col_name: "N",
        },
        {
          col_name: "O",
        },
        {
          col_name: "P",
        },
        {
          col_name: "Q",
        },
        {
          col_name: "R",
        },
        {
          col_name: "S",
        },
        {
          col_name: "T",
        },
        {
          col_name: "U",
        },
        {
          col_name: "V",
        },
        {
          col_name: "W",
        },
        {
          col_name: "X",
        },
      ],
      column_na: [
        {
          col_name: "AR",
        },
        {
          col_name: "AS",
        },
        {
          col_name: "AT",
        },
      ],

      row_month: [
        {
          row_name: "Jan",
        },
        {
          row_name: "Feb",
        },
        {
          row_name: "Mar",
        },
        {
          row_name: "Apr",
        },
        {
          row_name: "May",
        },
        {
          row_name: "Jun",
        },
        {
          row_name: "Jul",
        },
        {
          row_name: "Aug",
        },
        {
          row_name: "Sep",
        },
        {
          row_name: "Oct",
        },
        {
          row_name: "Nov",
        },
        {
          row_name: "Dec",
        },
      ],
      value_date_master: [
        {
          start_date: "01/01",
          end_date: "12/31",
        },
      ],
      value_date: [
        {
          month: "Jan",
          start_date: "01/01",
          end_date: "01/31",
        },

        {
          month: "Feb",
          start_date: "02/01",
          end_date: "02/28",
        },

        {
          month: "Mar",
          start_date: "03/01",
          end_date: "03/31",
        },

        {
          month: "Apr",
          start_date: "04/01",
          end_date: "04/30",
        },
        {
          month: "May",
          start_date: "05/01",
          end_date: "05/31",
        },
        {
          month: "Jun",
          start_date: "06/01",
          end_date: "06/30",
        },
        {
          month: "Jul",
          start_date: "07/01",
          end_date: "07/31",
        },
        {
          month: "Aug",
          start_date: "08/01",
          end_date: "08/30",
        },
        {
          month: "Sep",
          start_date: "09/01",
          end_date: "09/30",
        },
        {
          month: "Oct",
          start_date: "10/01",
          end_date: "10/31",
        },
        {
          month: "Nov",
          start_date: "11/01",
          end_date: "11/30",
        },
        {
          month: "Dec",
          start_date: "12/01",
          end_date: "12/31",
        },
      ],
      row_0_left: [
        {
          col_name: "L",
        },
        {
          col_name: "M",
        },
        {
          col_name: "N",
        },
        {
          col_name: "O",
        },
      ],
    };
  },
  mounted() {
    let login_status = this.$q.localStorage.getItem("login_status");

    if (login_status == null) {
      this.$router.push({
        path: "/",
      });
    }
  },
  methods: {
    logout() {
      this.basic = true;
    },
    confirm_log_out() {
      this.$router.push("/login");
      this.$q.localStorage.clear();
    },
    cancel_log_out() {
      this.basic = false;
    },
    async exportexcel() {
      if (this.org == "" || this.start == "" || this.end == "") {
        this.$q.notify({
          message: "Please Input Start Date End Date and Org",
          color: "red-9",
        });
      } else {
        this.$q.loading.show({
          spinner: QSpinnerGears,
          spinnerColor: "wthite",
          spinnerSize: 180,
          backgroundColor: "black",
        });
        this.total_pcs = [];
        this.total_yard = [];
        this.keep_cut_qty = [];
        this.keep_defect_1 = [];
        this.keep_defect_2 = [];
        this.keep_defect_3 = [];
        this.keep_defect_4 = [];
        this.keep_defect_5 = [];
        this.keep_defect_6 = [];
        this.keep_defect_7 = [];
        this.keep_defect_8 = [];
        this.keep_defect_9 = [];
        this.keep_defect_10 = [];
        this.sum_total_yard = [];
        this.recent_yard = [];
        this.recent_pcs = [];
        this.keep_total_yard = [];
        this.keep_total_per = [];
        this.keep_cut_qty_per = [];
        this.all_result_yard = [];
        this.all_result_yard2 = [];
        this.all_result_yard3 = [];
        this.all_result_yard4 = [];
        this.all_result_yard5 = [];
        this.all_result_yard6 = [];
        this.all_result_yard7 = [];
        this.all_result_yard8 = [];
        this.all_result_yard9 = [];
        this.all_result_yard10 = [];
        this.all_result_yard11 = [];
        this.all_result_pcs = [];
        this.all_result_pcs2 = [];
        this.all_result_pcs3 = [];
        this.all_result_pcs4 = [];
        this.all_result_pcs5 = [];
        this.all_result_pcs6 = [];
        this.all_result_pcs7 = [];
        this.all_result_pcs8 = [];
        this.all_result_pcs9 = [];
        this.all_result_pcs10 = [];
        this.all_result_pcs11 = [];
        this.row_result2 = [];
        this.row_result = [];
        this.rowexport_use = [];
        this.row_result_month = [];
        this.row_result_month_use = [];

        const workbook = new Excel.Workbook();
        workbook.creator = "Nanyang";
        workbook.lastModifiedBy = "Admin";
        workbook.created = new Date(2021, 8, 30);
        workbook.modified = new Date();
        workbook.lastPrinted = new Date(2021, 7, 27);

        const worksheet = workbook.addWorksheet("Data-h", {
          properties: { tabColor: { argb: "FF9966" } },
        });

        worksheet.columns = [
          { key: "A", width: 18 },
          { key: "B", width: 12 },
          { key: "C", width: 7.63 },
          { key: "D", width: 12 },
          { key: "E", width: 10 },
          { key: "F", width: 15 },
          { key: "G", width: 15 },
          { key: "H", width: 15 },
          { key: "I", width: 15 },
          { key: "J", width: 12 },
          { key: "K", width: 12 },
          { key: "L", width: 50 },
          { key: "M", width: 75 },
          { key: "N", width: 75 },
          { key: "O", width: 75 },
          { key: "P", width: 15 },
          { key: "Q", width: 6.7 },
          { key: "R", width: 18 },
          { key: "S", width: 6.7 },
          { key: "T", width: 13 },
          { key: "U", width: 6.7 },
          { key: "V", width: 6.7 },
          { key: "W", width: 13 },
          { key: "X", width: 6.7 },
          { key: "AC", width: 15.13 },
        ];

        //-----------------------------------------------------------------------------------------------------------------------------

        //sheet2

        console.log(this.row_result_item_code);
        const params3 = new FormData();

        params3.append("start", this.start_date);
        params3.append("end", this.end_date);
        params3.append("org", this.org);

        await axios({
          method: "post",
          url: this.$api_url + "/find_data.php/find_data_na",
          data: params3,
        }).then((resp) => {
          console.log(resp.data);
          if (resp.data.data.length > 0) {
            this.rowexport = [];
            resp.data.data.forEach((e) => {
              this.rowexport_use.push({
                shipment_date: e.SHIPMENT_DATE,
                so_no: e.SO_NO,
                so_no_doc: e.SO_NO_DOC,
                cust_name: e.CUST_NAME,
                style_ref: e.STYLE_REF,
                order_qty: e.ORDER_QTY,
                cut_qty: e.CUT_QTY,
                yard: e.YARD,
                item_code: e.ITEM_CODE,
                item_desc: e.ITEM_DESC,
                dept_code: e.DEPT_CODE,
                dept_name: e.DEPT_NAME,
                cpart_no1: e.CPART_NO1,
                cpart_no2: e.CPART_NO2,
                cpart_no3: e.CPART_NO3,
                reason: e.REASON,
                primary_quantity: e.PRIMARY_QUANTITY,
                org: e.ORG,
                g_yd: e.G_YD,
                yd_dz: e.YD_DZ,
                pcs: e.PCS,
              });
            });
          }
        });

        console.log(this.rowexport_use);
        worksheet.getCell("A1").value =
          "Re Cut Report" + "  " + this.org + " " + this.start + "-" + this.end;
        worksheet.mergeCells("A1:G1");

        worksheet.getCell("A2").value = "Shipment Date";
        worksheet.mergeCells("A2:A3");

        worksheet.getCell("B2").value = "S/O";
        worksheet.mergeCells("B2:B3");

        worksheet.getCell("C2").value = "Customer";
        worksheet.mergeCells("C2:C3");

        worksheet.getCell("D2").value = "Style";
        worksheet.mergeCells("D2:D3");

        worksheet.getCell("E2").value = "order";
        worksheet.mergeCells("E2:E3");

        worksheet.getCell("F2").value = "ยอดตัด";
        worksheet.mergeCells("F2:G2");

        worksheet.getCell("H2").value = "Y";
        worksheet.mergeCells("H2:H3");

        worksheet.getCell("I2").value = "KG";
        worksheet.mergeCells("I2:I3");

        worksheet.getCell("J2").value = "Yard per Dozen";
        worksheet.mergeCells("J2:J3");

        worksheet.getCell("K2").value = "Gram per Yard";
        worksheet.mergeCells("K2:K3");

        worksheet.getCell("L2").value = "ชนิดผ้า";
        worksheet.mergeCells("L2:L3");

        worksheet.getCell("M2").value = "Part";
        worksheet.mergeCells("M2:O3");

        worksheet.getCell("P2").value = "Re Cut Loss";
        worksheet.mergeCells("P2:R2");

        worksheet.getCell("S2").value = "";

        worksheet.getCell("T2").value = "เสียจากแผนกเย็บ";
        worksheet.mergeCells("T2:U2");

        worksheet.getCell("V2").value = "เสียจากผ้าตำหนิ";
        worksheet.mergeCells("V2:W2");

        worksheet.getCell("X2").value = "งานตัดคัดชิ้น";
        worksheet.mergeCells("X2:Y2");

        worksheet.getCell("Z2").value = "เสียจากแผนกตัด";
        worksheet.mergeCells("Z2:AA2");

        worksheet.getCell("AB2").value = "เสียจากแผนกรีด";
        worksheet.mergeCells("AB2:AC2");

        worksheet.getCell("AD2").value = "เสียจากแผนกแพด";
        worksheet.mergeCells("AD2:AE2");

        worksheet.getCell("AF2").value = "เสียจากแผนกพิมพ์";
        worksheet.mergeCells("AF2:AG2");

        worksheet.getCell("AH2").value = "เสียจากแผนกฟิวส์";
        worksheet.mergeCells("AH2:AI2");

        worksheet.getCell("AJ2").value = "เสียจากแผนกปัก";
        worksheet.mergeCells("AJ2:AK2");

        worksheet.getCell("AL2").value = "งานหาย";
        worksheet.mergeCells("AL2:AM2");

        worksheet.getCell("AN2").value = "Total";
        worksheet.mergeCells("AN2:AO2");

        worksheet.getCell("AR2").value = "NA";
        worksheet.mergeCells("AR2:AS2");

        worksheet.getCell("AT2").value = "Remark";
        worksheet.mergeCells("AT2:AT3");

        worksheet.getCell("F3").value = "Pcs.";
        worksheet.getCell("G3").value = "Yard";

        worksheet.getCell("P3").value = "Yards";
        worksheet.getCell("Q3").value = "Pcs.";
        worksheet.getCell("R3").value = "%";

        worksheet.getCell("T3").value = "Pcs";
        worksheet.getCell("U3").value = "Yard";
        worksheet.getCell("V3").value = "Pcs";
        worksheet.getCell("W3").value = "Yard";
        worksheet.getCell("X3").value = "Pcs";
        worksheet.getCell("Y3").value = "Yard";
        worksheet.getCell("Z3").value = "Pcs";
        worksheet.getCell("AA3").value = "Yard";
        worksheet.getCell("AB3").value = "Pcs";
        worksheet.getCell("AC3").value = "Yard";
        worksheet.getCell("AD3").value = "Pcs";
        worksheet.getCell("AE3").value = "Yard";
        worksheet.getCell("AF3").value = "Pcs";
        worksheet.getCell("AG3").value = "Yard";
        worksheet.getCell("AH3").value = "Pcs";
        worksheet.getCell("AI3").value = "Yard";
        worksheet.getCell("AJ3").value = "Pcs";
        worksheet.getCell("AK3").value = "Yard";
        worksheet.getCell("AL3").value = "Pcs";
        worksheet.getCell("AM3").value = "Yard";
        worksheet.getCell("AN3").value = "Pcs";
        worksheet.getCell("AO3").value = "Yard";
        worksheet.getCell("AR3").value = "Pcs";
        worksheet.getCell("AS3").value = "Yard";

        function numberWithCommas(val) {
          return val.toString().replace(/\B(?<!\.\d*)(?=(\d{3})+(?!\d))/g, ",");
        }

        for (var ax = 0; ax < this.rowexport_use.length; ax++) {
          worksheet.getCell("A" + [ax + 4]).value =
            this.rowexport_use[ax].shipment_date;

          if (this.rowexport_use[ax].so_no.length == 3) {
            this.rowexport_use[ax].so_no = [0] + this.rowexport_use[ax].so_no;
          } else if (this.rowexport_use[ax].so_no.length == 2) {
            this.rowexport_use[ax].so_no =
              [0] + [0] + this.rowexport_use[ax].so_no;
          } else if (this.rowexport_use[ax].so_no.length == 1) {
            this.rowexport_use[ax].so_no =
              [0] + [0] + [0] + this.rowexport_use[ax].so_no;
          }

          worksheet.getCell("B" + [ax + 4]).value =
            this.rowexport_use[ax].so_no;

          worksheet.getCell("C" + [ax + 4]).value =
            this.rowexport_use[ax].cust_name;

          worksheet.getCell("D" + [ax + 4]).value =
            this.rowexport_use[ax].style_ref;
          /*  if (this.rowexport_use[ax].order_qty > 999) {
            this.rowexport_use[ax].order_qty = numberWithCommas(
              this.rowexport_use[ax].order_qty
            );
          } */

          worksheet.getCell("E" + [ax + 4]).value = Number(
            this.rowexport_use[ax].order_qty
          );
          worksheet.getCell("E" + [ax + 4]).numFmt = "#,##0_);[Red](#,##0)";

          worksheet.getCell("F" + [ax + 4]).value = Number(
            this.rowexport_use[ax].cut_qty
          );
          worksheet.getCell("F" + [ax + 4]).numFmt = "#,##0_);[Red](#,##0)";

          worksheet.getCell("G" + [ax + 4]).value = Number(
            this.rowexport_use[ax].yard
          );
          worksheet.getCell("G" + [ax + 4]).numFmt = "#,##0.00_);(#,##0.00)";

          worksheet.getCell("I" + [ax + 4]).value = "-";

          worksheet.getCell("J" + [ax + 4]).value = Number(
            this.rowexport_use[ax].yd_dz
          );
          worksheet.getCell("J" + [ax + 4]).numFmt = "0.00";

          worksheet.getCell("K" + [ax + 4]).value = Number(
            this.rowexport_use[ax].g_yd
          );
          worksheet.getCell("K" + [ax + 4]).numFmt = "0.00";

          worksheet.getCell("L" + [ax + 4]).value =
            this.rowexport_use[ax].item_desc;

          worksheet.getCell("M" + [ax + 4]).value =
            this.rowexport_use[ax].cpart_no1;

          worksheet.getCell("N" + [ax + 4]).value =
            this.rowexport_use[ax].cpart_no2;

          worksheet.getCell("O" + [ax + 4]).value =
            this.rowexport_use[ax].cpart_no3;

          worksheet.getCell("P " + [ax + 4]).value = "";

          if (this.rowexport_use[ax].reason == "เสียจากแผนกเย็บ") {
            worksheet.getCell("T " + [ax + 4]).value = Number(
              this.rowexport_use[ax].pcs
            );
            worksheet.getCell("T " + [ax + 4]).numFmt = "0.0";
            worksheet.getCell("U " + [ax + 4]).value = Number(
              this.rowexport_use[ax].primary_quantity
            );
            worksheet.getCell("U " + [ax + 4]).numFmt = "0.0";

            if (this.rowexport_use[ax].primary_quantity > 0) {
              this.all_result_yard.push({
                value: this.rowexport_use[ax].primary_quantity,
              });
            } else {
              this.all_result_yard.push({
                value: 0,
              });
            }

            if (this.rowexport_use[ax].pcs > 0) {
              this.all_result_pcs.push({
                value: this.rowexport_use[ax].pcs,
              });
            } else {
              this.all_result_pcs.push({
                value: 0,
              });
            }
          }

          if (this.rowexport_use[ax].reason == "เสียจากผ้าตำหนิ") {
            worksheet.getCell("V " + [ax + 4]).value = Number(
              this.rowexport_use[ax].pcs
            );
            worksheet.getCell("V " + [ax + 4]).numFmt = "0.0";
            worksheet.getCell("W " + [ax + 4]).value = Number(
              this.rowexport_use[ax].primary_quantity
            );
            worksheet.getCell("W " + [ax + 4]).numFmt = "0.0";
            if (this.rowexport_use[ax].primary_quantity > 0) {
              this.all_result_yard2.push({
                value: this.rowexport_use[ax].primary_quantity,
              });
            } else {
              this.all_result_yard2.push({
                value: 0,
              });
            }

            if (this.rowexport_use[ax].pcs > 0) {
              this.all_result_pcs2.push({
                value: this.rowexport_use[ax].pcs,
              });
            } else {
              this.all_result_pcs2.push({
                value: 0,
              });
            }
          }

          if (this.rowexport_use[ax].reason == "งานตัดคัดชิ้น") {
            worksheet.getCell("X " + [ax + 4]).value = Number(
              this.rowexport_use[ax].pcs
            );
            worksheet.getCell("X " + [ax + 4]).numFmt = "0.0";
            worksheet.getCell("Y " + [ax + 4]).value = Number(
              this.rowexport_use[ax].primary_quantity
            );
            console.log(this.rowexport_use[ax].primary_quantity);
            worksheet.getCell("Y " + [ax + 4]).numFmt = "0.0";
            if (this.rowexport_use[ax].primary_quantity > 0) {
              this.all_result_yard3.push({
                value: this.rowexport_use[ax].primary_quantity,
              });
            } else {
              this.all_result_yard3.push({
                value: 0,
              });
            }

            if (this.rowexport_use[ax].pcs > 0) {
              this.all_result_pcs3.push({
                value: this.rowexport_use[ax].pcs,
              });
            } else {
              this.all_result_pcs3.push({
                value: 0,
              });
            }
          }
          if (this.rowexport_use[ax].reason == "เสียจากแผนกตัด") {
            worksheet.getCell("Z " + [ax + 4]).value = Number(
              this.rowexport_use[ax].pcs
            );
            worksheet.getCell("Z " + [ax + 4]).numFmt = "0.0";
            worksheet.getCell("AA " + [ax + 4]).value = Number(
              this.rowexport_use[ax].primary_quantity
            );
            worksheet.getCell("AA " + [ax + 4]).numFmt = "0.0";
            if (this.rowexport_use[ax].primary_quantity > 0) {
              this.all_result_yard4.push({
                value: this.rowexport_use[ax].primary_quantity,
              });
            } else {
              this.all_result_yard4.push({
                value: 0,
              });
            }

            if (this.rowexport_use[ax].pcs > 0) {
              this.all_result_pcs4.push({
                value: this.rowexport_use[ax].pcs,
              });
            } else {
              this.all_result_pcs4.push({
                value: 0,
              });
            }
          }

          if (this.rowexport_use[ax].reason == "เสียจากแผนกรีด") {
            worksheet.getCell("AB " + [ax + 4]).value = Number(
              this.rowexport_use[ax].pcs
            );
            worksheet.getCell("AB " + [ax + 4]).numFmt = "0.0";
            worksheet.getCell("AC " + [ax + 4]).value = Number(
              this.rowexport_use[ax].primary_quantity
            );
            worksheet.getCell("AC " + [ax + 4]).numFmt = "0.0";
            if (this.rowexport_use[ax].primary_quantity > 0) {
              this.all_result_yard5.push({
                value: this.rowexport_use[ax].primary_quantity,
              });
            } else {
              this.all_result_yard5.push({
                value: 0,
              });
            }

            if (this.rowexport_use[ax].pcs > 0) {
              this.all_result_pcs5.push({
                value: this.rowexport_use[ax].pcs,
              });
            } else {
              this.all_result_pcs5.push({
                value: 0,
              });
            }
          }

          if (this.rowexport_use[ax].reason == "เสียจากแผนกแพด") {
            worksheet.getCell("AD " + [ax + 4]).value = Number(
              this.rowexport_use[ax].pcs
            );
            worksheet.getCell("AD " + [ax + 4]).numFmt = "0.0";
            worksheet.getCell("AE " + [ax + 4]).value = Number(
              this.rowexport_use[ax].primary_quantity
            );
            worksheet.getCell("AE " + [ax + 4]).numFmt = "0.0";
            if (this.rowexport_use[ax].primary_quantity > 0) {
              this.all_result_yard6.push({
                value: this.rowexport_use[ax].primary_quantity,
              });
            } else {
              this.all_result_yard6.push({
                value: 0,
              });
            }

            if (this.rowexport_use[ax].pcs > 0) {
              this.all_result_pcs6.push({
                value: this.rowexport_use[ax].pcs,
              });
            } else {
              this.all_result_pcs6.push({
                value: 0,
              });
            }
          }
          if (this.rowexport_use[ax].reason == "เสียจากแผนกพิมพ์") {
            worksheet.getCell("AF " + [ax + 4]).value = Number(
              this.rowexport_use[ax].pcs
            );
            worksheet.getCell("AF " + [ax + 4]).numFmt = "0.0";
            worksheet.getCell("AG " + [ax + 4]).value = Number(
              this.rowexport_use[ax].primary_quantity
            );
            worksheet.getCell("AG " + [ax + 4]).numFmt = "0.0";
            if (this.rowexport_use[ax].primary_quantity > 0) {
              this.all_result_yard7.push({
                value: this.rowexport_use[ax].primary_quantity,
              });
            } else {
              this.all_result_yard7.push({
                value: 0,
              });
            }

            if (this.rowexport_use[ax].pcs > 0) {
              this.all_result_pcs7.push({
                value: this.rowexport_use[ax].pcs,
              });
            } else {
              this.all_result_pcs7.push({
                value: 0,
              });
            }
          }
          if (this.rowexport_use[ax].reason == "เสียจากแผนกฟิวส์") {
            worksheet.getCell("AH " + [ax + 4]).value = Number(
              this.rowexport_use[ax].pcs
            );
            worksheet.getCell("AH " + [ax + 4]).numFmt = "0.0";
            worksheet.getCell("AI " + [ax + 4]).value = Number(
              this.rowexport_use[ax].primary_quantity
            );
            worksheet.getCell("AI " + [ax + 4]).numFmt = "0.0";
            if (this.rowexport_use[ax].primary_quantity > 0) {
              this.all_result_yard9.push({
                value: this.rowexport_use[ax].primary_quantity,
              });
            } else {
              this.all_result_yard9.push({
                value: 0,
              });
            }

            if (this.rowexport_use[ax].pcs > 0) {
              this.all_result_pcs9.push({
                value: this.rowexport_use[ax].pcs,
              });
            } else {
              this.all_result_pcs9.push({
                value: 0,
              });
            }
          }

          if (this.rowexport_use[ax].reason == "เสียจากแผนกปัก") {
            worksheet.getCell("AJ " + [ax + 4]).value = Number(
              this.rowexport_use[ax].pcs
            );
            worksheet.getCell("AJ " + [ax + 4]).numFmt = "0.0";
            worksheet.getCell("AK " + [ax + 4]).value = Number(
              this.rowexport_use[ax].primary_quantity
            );
            worksheet.getCell("AK " + [ax + 4]).numFmt = "0.0";
            this.all_result_yard9.push({
              value: this.rowexport_use[ax].primary_quantity,
            });

            this.all_result_pcs9.push({
              value: this.rowexport_use[ax].pcs,
            });
          }
          if (this.rowexport_use[ax].reason == "งานหาย") {
            worksheet.getCell("AL " + [ax + 4]).value = Number(
              this.rowexport_use[ax].pcs
            );
            worksheet.getCell("AL " + [ax + 4]).numFmt = "0.0";
            worksheet.getCell("AM " + [ax + 4]).value = Number(
              this.rowexport_use[ax].primary_quantity
            );
            worksheet.getCell("AM " + [ax + 4]).numFmt = "0.0";
            if (this.rowexport_use[ax].primary_quantity > 0) {
              this.all_result_yard10.push({
                value: this.rowexport_use[ax].primary_quantity,
              });
            } else {
              this.all_result_yard10.push({
                value: 0,
              });
            }

            if (this.rowexport_use[ax].pcs > 0) {
              this.all_result_pcs10.push({
                value: this.rowexport_use[ax].pcs,
              });
            } else {
              this.all_result_pcs10.push({
                value: 0,
              });
            }
          }
          if (this.rowexport_use[ax].reason == "NA") {
            worksheet.getCell("AR " + [ax + 4]).value = Number(
              this.rowexport_use[ax].pcs
            );
            worksheet.getCell("AR " + [ax + 4]).numFmt = "0.0";
            worksheet.getCell("AS " + [ax + 4]).value = Number(
              this.rowexport_use[ax].primary_quantity
            );
            worksheet.getCell("AR " + [ax + 4]).numFmt = "0.0";
            if (this.rowexport_use[ax].primary_quantity > 0) {
              this.all_result_yard11.push({
                value: this.rowexport_use[ax].primary_quantity,
              });
            } else {
              this.all_result_yard11.push({
                value: 0,
              });
            }

            if (this.rowexport_use[ax].pcs > 0) {
              this.all_result_pcs11.push({
                value: this.rowexport_use[ax].pcs,
              });
            } else {
              this.all_result_pcs11.push({
                value: 0,
              });
            }
          }
        }

        for (var ax = 0; ax < this.rowexport_use.length; ax++) {
          worksheet.getCell("P " + [ax + 4]).value =
            this.rowexport_use[ax].primary_quantity;
          worksheet.getCell("H " + [ax + 4]).value =
            this.rowexport_use[ax].primary_quantity;
          worksheet.getCell("Q " + [ax + 4]).value = this.rowexport_use[ax].pcs;

          if (this.rowexport_use[ax].pcs != null) {
            worksheet.getCell("AN " + [ax + 4]).value = Number(
              this.rowexport_use[ax].pcs
            );
            worksheet.getCell("AN " + [ax + 4]).numFmt = "0.00";
          } else {
            worksheet.getCell("AN " + [ax + 4]).value = 0;
          }

          if (this.rowexport_use[ax].primary_quantity != null) {
            worksheet.getCell("AO " + [ax + 4]).value = Number(
              this.rowexport_use[ax].primary_quantity
            );
            worksheet.getCell("AO " + [ax + 4]).numFmt = "0.00";
          } else {
            worksheet.getCell("AO " + [ax + 4]).value = 0;
          }

          this.recent_loss_per = 0;
          this.recent_loss_per =
            (this.rowexport_use[ax].primary_quantity /
              this.rowexport_use[ax].yard) *
            100;

          if (isNaN(this.recent_loss_per) == true) {
            this.recent_loss_per = 0;
          }

          if (isFinite(this.recent_loss_per) == false) {
            this.recent_loss_per = 0;
          }

          if (this.recent_loss_per > 0) {
            worksheet.getCell("R " + [ax + 4]).value = "";
            worksheet.getCell("R " + [ax + 4]).numFmt = "0.00%";
          } else {
            worksheet.getCell("R " + [ax + 4]).value = "";
            worksheet.getCell("R " + [ax + 4]).numFmt = "0.00%";
          }
        }
        var ac = 0;

        this.number_total = ac + 4 + this.rowexport_use.length;
        this.total_sum_yardes = 0;
        this.total_sum_pcs_bot = 0;
        this.total_sum_yard_bot = 0;
        for (var ax = 0; ax < this.rowexport_use.length; ax++) {
          this.total_sum_yardes =
            Number(this.total_sum_yardes) + Number(this.rowexport_use[ax].yard);

          this.total_sum_yard_bot =
            Number(this.total_sum_yard_bot) +
            Number(this.rowexport_use[ax].primary_quantity);

          this.total_sum_pcs_bot =
            Number(this.total_sum_pcs_bot) + Number(this.rowexport_use[ax].pcs);
        }

        worksheet.getCell("A" + this.number_total).value = "Total";
        worksheet.mergeCells(
          "A" + this.number_total + ": F" + this.number_total
        );
        worksheet.getCell("G" + this.number_total).value = "";
        worksheet.getCell("G" + this.number_total).numFmt =
          "#,##0.00_);(#,##0.00)";
        console.log(this.all_result_yard);
        console.log(this.all_result_yard2);
        console.log(this.all_result_yard3);

        this.total_sum_yard = 0;
        this.total_sum_yard2 = 0;
        this.total_sum_yard3 = 0;
        this.total_sum_yard4 = 0;
        this.total_sum_yard5 = 0;
        this.total_sum_yard6 = 0;
        this.total_sum_yard7 = 0;
        this.total_sum_yard8 = 0;
        this.total_sum_yard9 = 0;
        this.total_sum_yard10 = 0;
        this.total_sum_yard11 = 0;
        for (var ax = 0; ax < this.all_result_yard.length; ax++) {
          this.total_sum_yard =
            Number(this.total_sum_yard) +
            Number(this.all_result_yard[ax].value);
        }
        for (var ax = 0; ax < this.all_result_yard2.length; ax++) {
          this.total_sum_yard2 =
            Number(this.total_sum_yard2) +
            Number(this.all_result_yard2[ax].value);
        }

        for (var ax = 0; ax < this.all_result_yard3.length; ax++) {
          this.total_sum_yard3 =
            Number(this.total_sum_yard3) +
            Number(this.all_result_yard3[ax].value);
        }
        for (var ax = 0; ax < this.all_result_yard4.length; ax++) {
          this.total_sum_yard4 =
            Number(this.total_sum_yard4) +
            Number(this.all_result_yard4[ax].value);
        }
        for (var ax = 0; ax < this.all_result_yard5.length; ax++) {
          this.total_sum_yard5 =
            Number(this.total_sum_yard5) +
            Number(this.all_result_yard5[ax].value);
        }
        for (var ax = 0; ax < this.all_result_yard6.length; ax++) {
          this.total_sum_yard6 =
            Number(this.total_sum_yard6) +
            Number(this.all_result_yard6[ax].value);
        }
        for (var ax = 0; ax < this.all_result_yard7.length; ax++) {
          this.total_sum_yard7 =
            Number(this.total_sum_yard7) +
            Number(this.all_result_yard7[ax].value);
        }
        for (var ax = 0; ax < this.all_result_yard8.length; ax++) {
          this.total_sum_yard8 =
            Number(this.total_sum_yard8) +
            Number(this.all_result_yard8[ax].value);
        }
        for (var ax = 0; ax < this.all_result_yard9.length; ax++) {
          this.total_sum_yard9 =
            Number(this.total_sum_yard9) +
            Number(this.all_result_yard9[ax].value);
        }
        for (var ax = 0; ax < this.all_result_yard10.length; ax++) {
          this.total_sum_yard10 =
            Number(this.total_sum_yard10) +
            Number(this.all_result_yard10[ax].value);
        }
        for (var ax = 0; ax < this.all_result_yard11.length; ax++) {
          this.total_sum_yard11 =
            Number(this.total_sum_yard11) +
            Number(this.all_result_yard11[ax].value);
        }
        if (
          this.total_sum_yard == 0 ||
          this.total_sum_yard == undefined ||
          isNaN(this.total_sum_yard) == true
        ) {
          this.total_sum_yard = 0;
        }
        if (
          this.total_sum_yard2 == 0 ||
          this.total_sum_yard2 == undefined ||
          isNaN(this.total_sum_yard2) == true
        ) {
          this.total_sum_yard2 = 0;
        }
        if (
          this.total_sum_yard3 == 0 ||
          this.total_sum_yard3 == undefined ||
          isNaN(this.total_sum_yard3) == true
        ) {
          this.total_sum_yard3 = 0;
        }
        if (
          this.total_sum_yard4 == 0 ||
          this.total_sum_yard4 == undefined ||
          isNaN(this.total_sum_yard4) == true
        ) {
          this.total_sum_yard4 = 0;
        }
        if (
          this.total_sum_yard5 == 0 ||
          this.total_sum_yard5 == undefined ||
          isNaN(this.total_sum_yard5) == true
        ) {
          this.total_sum_yard5 = 0;
        }
        if (
          this.total_sum_yard6 == 0 ||
          this.total_sum_yard6 == undefined ||
          isNaN(this.total_sum_yard6) == true
        ) {
          this.total_sum_yard6 = 0;
        }

        if (
          this.total_sum_yard7 == 0 ||
          this.total_sum_yard7 == undefined ||
          isNaN(this.total_sum_yard7) == true
        ) {
          this.total_sum_yard7 = 0;
        }

        if (
          this.total_sum_yard7 == 0 ||
          this.total_sum_yard7 == undefined ||
          isNaN(this.total_sum_yard7) == true
        ) {
          this.total_sum_yard7 = 0;
        }

        if (
          this.total_sum_yard8 == 0 ||
          this.total_sum_yard8 == undefined ||
          isNaN(this.total_sum_yard8) == true
        ) {
          this.total_sum_yard8 = 0;
        }

        if (
          this.total_sum_yard9 == 0 ||
          this.total_sum_yard9 == undefined ||
          isNaN(this.total_sum_yard9) == true
        ) {
          this.total_sum_yard9 = 0;
        }

        if (
          this.total_sum_yard10 == 0 ||
          this.total_sum_yard10 == undefined ||
          isNaN(this.total_sum_yard10) == true
        ) {
          this.total_sum_yard10 = 0;
        }
        if (
          this.total_sum_yard11 == 0 ||
          this.total_sum_yard11 == undefined ||
          isNaN(this.total_sum_yard11) == true
        ) {
          this.total_sum_yard11 = 0;
        }

        this.total_sum_pcs = 0;
        this.total_sum_pcs2 = 0;
        this.total_sum_pcs3 = 0;
        this.total_sum_pcs4 = 0;
        this.total_sum_pcs5 = 0;
        this.total_sum_pcs6 = 0;
        this.total_sum_pcs7 = 0;
        this.total_sum_pcs8 = 0;
        this.total_sum_pcs9 = 0;
        this.total_sum_pcs10 = 0;
        this.total_sum_pcs11 = 0;
        for (var ax = 0; ax < this.all_result_pcs.length; ax++) {
          this.total_sum_pcs =
            Number(this.total_sum_pcs) + Number(this.all_result_pcs[ax].value);
        }
        for (var ax = 0; ax < this.all_result_pcs2.length; ax++) {
          this.total_sum_pcs2 =
            Number(this.total_sum_pcs2) +
            Number(this.all_result_pcs2[ax].value);
        }
        for (var ax = 0; ax < this.all_result_pcs3.length; ax++) {
          this.total_sum_pcs3 =
            Number(this.total_sum_pcs3) +
            Number(this.all_result_pcs3[ax].value);
        }

        for (var ax = 0; ax < this.all_result_pcs4.length; ax++) {
          this.total_sum_pcs4 =
            Number(this.total_sum_pcs4) +
            Number(this.all_result_pcs4[ax].value);
        }

        for (var ax = 0; ax < this.all_result_pcs5.length; ax++) {
          this.total_sum_pcs5 =
            Number(this.total_sum_pcs5) +
            Number(this.all_result_pcs5[ax].value);
        }

        for (var ax = 0; ax < this.all_result_pcs6.length; ax++) {
          this.total_sum_pcs6 =
            Number(this.total_sum_pcs6) +
            Number(this.all_result_pcs6[ax].value);
        }

        for (var ax = 0; ax < this.all_result_pcs7.length; ax++) {
          this.total_sum_pcs7 =
            Number(this.total_sum_pcs7) +
            Number(this.all_result_pcs7[ax].value);
        }

        for (var ax = 0; ax < this.all_result_pcs8.length; ax++) {
          this.total_sum_pcs8 =
            Number(this.total_sum_pcs8) +
            Number(this.all_result_pcs8[ax].value);
        }

        for (var ax = 0; ax < this.all_result_pcs9.length; ax++) {
          this.total_sum_pcs9 =
            Number(this.total_sum_pcs9) +
            Number(this.all_result_pcs9[ax].value);
        }

        for (var ax = 0; ax < this.all_result_pcs10.length; ax++) {
          this.total_sum_pcs10 =
            Number(this.total_sum_pcs10) +
            Number(this.all_result_pcs10[ax].value);
        }

        for (var ax = 0; ax < this.all_result_pcs11.length; ax++) {
          this.total_sum_pcs11 =
            Number(this.total_sum_pcs11) +
            Number(this.all_result_pcs11[ax].value);
        }

        if (
          this.total_sum_pcs == 0 ||
          this.total_sum_pcs == undefined ||
          isNaN(this.total_sum_pcs) == true
        ) {
          this.total_sum_pcs = 0;
        }
        if (
          this.total_sum_pcs2 == 0 ||
          this.total_sum_pcs2 == undefined ||
          isNaN(this.total_sum_pcs2) == true
        ) {
          this.total_sum_pcs2 = 0;
        }
        if (
          this.total_sum_pcs3 == 0 ||
          this.total_sum_pcs3 == undefined ||
          isNaN(this.total_sum_pcs3) == true
        ) {
          this.total_sum_pcs3 = 0;
        }
        if (
          this.total_sum_pcs4 == 0 ||
          this.total_sum_pcs4 == undefined ||
          isNaN(this.total_sum_pcs4) == true
        ) {
          this.total_sum_pcs4 = 0;
        }
        if (
          this.total_sum_pcs5 == 0 ||
          this.total_sum_pcs5 == undefined ||
          isNaN(this.total_sum_pcs5) == true
        ) {
          this.total_sum_pcs5 = 0;
        }
        if (
          this.total_sum_pcs6 == 0 ||
          this.total_sum_pcs6 == undefined ||
          isNaN(this.total_sum_pcs6) == true
        ) {
          this.total_sum_pcs6 = 0;
        }

        if (
          this.total_sum_pcs7 == 0 ||
          this.total_sum_pcs7 == undefined ||
          isNaN(this.total_sum_pcs7) == true
        ) {
          this.total_sum_pcs7 = 0;
        }

        if (
          this.total_sum_pcs7 == 0 ||
          this.total_sum_pcs7 == undefined ||
          isNaN(this.total_sum_pcs7) == true
        ) {
          this.total_sum_pcs7 = 0;
        }

        if (
          this.total_sum_pcs8 == 0 ||
          this.total_sum_pcs8 == undefined ||
          isNaN(this.total_sum_pcs8) == true
        ) {
          this.total_sum_pcs8 = 0;
        }

        if (
          this.total_sum_pcs9 == 0 ||
          this.total_sum_pcs9 == undefined ||
          isNaN(this.total_sum_pcs9) == true
        ) {
          this.total_sum_pcs9 = 0;
        }

        if (
          this.total_sum_pcs10 == 0 ||
          this.total_sum_pcs10 == undefined ||
          isNaN(this.total_sum_pcs10) == true
        ) {
          this.total_sum_pcs10 = 0;
        }
        if (
          this.total_sum_pcs11 == 0 ||
          this.total_sum_pcs11 == undefined ||
          isNaN(this.total_sum_pcs11) == true
        ) {
          this.total_sum_pcs11 = 0;
        }

        worksheet.getCell("T" + this.number_total).value = this.total_sum_pcs;
        worksheet.getCell("U" + this.number_total).value = this.total_sum_yard;

        worksheet.getCell("V" + this.number_total).value = this.total_sum_pcs2;
        worksheet.getCell("W" + this.number_total).value = this.total_sum_yard2;

        worksheet.getCell("X" + this.number_total).value = this.total_sum_pcs3;
        worksheet.getCell("Y" + this.number_total).value = this.total_sum_yard3;

        worksheet.getCell("Z" + this.number_total).value = this.total_sum_pcs4;
        worksheet.getCell("AA" + this.number_total).value =
          this.total_sum_yard4;

        worksheet.getCell("AB" + this.number_total).value = this.total_sum_pcs5;
        worksheet.getCell("AC" + this.number_total).value =
          this.total_sum_yard5;

        worksheet.getCell("AD" + this.number_total).value = this.total_sum_pcs6;
        worksheet.getCell("AE" + this.number_total).value =
          this.total_sum_yard6;

        worksheet.getCell("AF" + this.number_total).value = this.total_sum_pcs7;
        worksheet.getCell("AG" + this.number_total).value =
          this.total_sum_yard7;

        worksheet.getCell("AH" + this.number_total).value = this.total_sum_pcs8;
        worksheet.getCell("AI" + this.number_total).value =
          this.total_sum_yard8;

        worksheet.getCell("AJ" + this.number_total).value = this.total_sum_pcs9;
        worksheet.getCell("AK" + this.number_total).value =
          this.total_sum_yard9;

        worksheet.getCell("AL" + this.number_total).value =
          this.total_sum_pcs10;
        worksheet.getCell("AM" + this.number_total).value =
          this.total_sum_yard10;

        worksheet.getCell("AR" + this.number_total).value =
          this.total_sum_pcs11;
        worksheet.getCell("AS" + this.number_total).value =
          this.total_sum_yard11;

        this.sum_yard_bot = 0;
        this.sum_yard_bot =
          Number(this.sum_yard_bot) +
          Number(this.total_sum_yard) +
          Number(this.total_sum_yard2) +
          Number(this.total_sum_yard3) +
          Number(this.total_sum_yard4) +
          Number(this.total_sum_yard5) +
          Number(this.total_sum_yard6) +
          Number(this.total_sum_yard7) +
          Number(this.total_sum_yard8) +
          Number(this.total_sum_yard9) +
          Number(this.total_sum_yard10) +
          Number(this.total_sum_yard11);

        this.sum_pcs_bot = 0;
        this.sum_pcs_bot =
          Number(this.sum_pcs_bot) +
          Number(this.total_sum_pcs) +
          Number(this.total_sum_pcs2) +
          Number(this.total_sum_pcs3) +
          Number(this.total_sum_pcs4) +
          Number(this.total_sum_pcs5) +
          Number(this.total_sum_pcs6) +
          Number(this.total_sum_pcs7) +
          Number(this.total_sum_pcs8) +
          Number(this.total_sum_pcs9) +
          Number(this.total_sum_pcs10) +
          Number(this.total_sum_pcs11);

        if (this.sum_yard_bot > 0 || isNaN(this.sum_yard_bot) == false) {
          worksheet.getCell("AO" + this.number_total).value = this.sum_yard_bot;
        } else {
          worksheet.getCell("AO" + this.number_total).value = 0;
        }

        if (this.sum_pcs_bot > 0 || isNaN(this.sum_pcs_bot) == false) {
          worksheet.getCell("AN" + this.number_total).value = this.sum_pcs_bot;
        } else {
          worksheet.getCell("AN" + this.number_total).value = 0;
        }

        this.recent_per = 0;
        this.recent_per = (this.sum_yard_bot / this.total_sum_yardes) * 100;
        worksheet.getCell("R" + this.number_total).value = "" / 100;
        worksheet.getCell("R" + this.number_total).numFmt = "0.00%";

        //sheet3
        //control grid
        for (var ax = 0; ax < this.column_main.length; ax++) {
          for (var bz = 1; bz < this.number_total; bz++) {
            worksheet.getCell(this.column_main[ax].col_name + bz).font = {
              name: "Angsana New",
              color: { argb: "FF000000" },
              family: 4,
              size: 14,
              bold: false,
            };

            //เดือน
            worksheet.getCell(this.column_main[ax].col_name + bz).alignment = {
              horizontal: "center",
              vertical: "middle",
            };

            worksheet.getCell(this.column_main[ax].col_name + bz).border = {
              top: { style: "thin" },
              left: { style: "thin" },
              bottom: { style: "thin" },
              right: { style: "thin" },
            };
          }
        }

        for (var ax = 0; ax < this.column_na.length; ax++) {
          for (var bz = 1; bz < this.number_total + 1; bz++) {
            worksheet.getCell(this.column_na[ax].col_name + bz).font = {
              name: "Angsana New",
              color: { argb: "FF000000" },
              family: 4,
              size: 14,
              bold: false,
            };

            //เดือน
            worksheet.getCell(this.column_na[ax].col_name + bz).alignment = {
              horizontal: "center",
              vertical: "middle",
            };

            worksheet.getCell(this.column_na[ax].col_name + bz).border = {
              top: { style: "thin" },
              left: { style: "thin" },
              bottom: { style: "thin" },
              right: { style: "thin" },
            };
          }
        }

        for (var ax = 0; ax < this.column_main.length; ax++) {
          worksheet.getCell(
            this.column_main[ax].col_name + this.number_total
          ).font = {
            name: "Angsana New",
            color: { argb: "FF000000" },
            family: 4,
            size: 14,
            bold: false,
          };

          //เดือน
          worksheet.getCell(
            this.column_main[ax].col_name + this.number_total
          ).alignment = {
            horizontal: "center",
            vertical: "middle",
          };

          worksheet.getCell(
            this.column_main[ax].col_name + this.number_total
          ).border = {
            top: { style: "double" },
            left: { style: "double" },
            bottom: { style: "double" },
            right: { style: "double" },
          };
        }

        //color special grid
        for (var ax = 0; ax < this.rowexport_use.length; ax++) {
          worksheet.getCell("G" + [ax + 4]).fill = {
            type: "pattern",
            pattern: "solid",
            fgColor: { argb: "FFAEE395" },
            bgColor: { argb: "FFAEE395" },
          };

          worksheet.getCell("P" + [ax + 4]).fill = {
            type: "pattern",
            pattern: "solid",
            fgColor: { argb: "FFAEE395" },
            bgColor: { argb: "FFAEE395" },
          };

          worksheet.getCell("Q" + [ax + 4]).fill = {
            type: "pattern",
            pattern: "solid",
            fgColor: { argb: "FF95CCE3" },
            bgColor: { argb: "FF95CCE3" },
          };

          worksheet.getCell("R" + [ax + 4]).fill = {
            type: "pattern",
            pattern: "solid",
            fgColor: { argb: "FFEAB1C6" },
            bgColor: { argb: "FFEAB1C6" },
          };

          worksheet.getCell("R" + [ax + 4]).fill = {
            type: "pattern",
            pattern: "solid",
            fgColor: { argb: "FFEAB1C6" },
            bgColor: { argb: "FFEAB1C6" },
          };

          worksheet.getCell("AN" + [ax + 4]).fill = {
            type: "pattern",
            pattern: "solid",
            fgColor: { argb: "FFEDEAB4" },
            bgColor: { argb: "FFEDEAB4" },
          };

          worksheet.getCell("S" + [ax + 4]).fill = {
            type: "pattern",
            pattern: "solid",
            fgColor: { argb: "FF571300" },
            bgColor: { argb: "FF571300" },
          };

          worksheet.getCell("AO" + [ax + 4]).fill = {
            type: "pattern",
            pattern: "solid",
            fgColor: { argb: "FFEDEAB4" },
            bgColor: { argb: "FFEDEAB4" },
          };
        }

        worksheet.getCell("S" + [this.number_total]).fill = {
          type: "pattern",
          pattern: "solid",
          fgColor: { argb: "FF571300" },
          bgColor: { argb: "FF571300" },
        };
        worksheet.getCell("A1").font = {
          name: "Angsana New",
          color: { argb: "FF376DE2" },
          family: 4,
          size: 20,
          bold: false,
        };

        worksheet.getCell("A1").alignment = {
          horizontal: "left",
          vertical: "middle",
        };

        worksheet.getCell("A1").border = {
          top: { style: "thin" },
          left: { style: "thin" },
          bottom: { style: "thin" },
          right: { style: "thin" },
        };

        if (this.monthx == "01") {
          this.monthx = "JAN";
        }
        if (this.monthx == "02") {
          this.monthx = "FEB";
        }
        if (this.monthx == "03") {
          this.monthx = "MAR";
        }
        if (this.monthx == "04") {
          this.monthx = "APR";
        }
        if (this.monthx == "05") {
          this.monthx = "MAY";
        }
        if (this.monthx == "06") {
          this.monthx = "JUN";
        }
        if (this.monthx == "07") {
          this.monthx = "JUL";
        }
        if (this.monthx == "08") {
          this.monthx = "AUG";
        }
        if (this.monthx == "09") {
          this.monthx = "SEP";
        }
        if (this.monthx == "10") {
          this.monthx = "OCT";
        }
        if (this.monthx == "11") {
          this.monthx = "NOV";
        }
        if (this.monthx == "12") {
          this.monthx = "DEC";
        }

        for (var ax = 0; ax < this.row_0_left.length; ax++) {
          for (var bz = 1; bz < this.number_total; bz++) {
            worksheet.getCell(this.row_0_left[ax].col_name + bz).font = {
              name: "Angsana New",
              color: { argb: "FF000000" },
              family: 4,
              size: 14,
              bold: false,
            };

            //เดือน
            worksheet.getCell(this.row_0_left[ax].col_name + bz).alignment = {
              horizontal: "left",
              vertical: "middle",
            };

            worksheet.getCell(this.row_0_left[ax].col_name + bz).border = {
              top: { style: "thin" },
              left: { style: "thin" },
              bottom: { style: "thin" },
              right: { style: "thin" },
            };
          }
        }
        worksheet.getCell("L2").alignment = {
          horizontal: "center",
          vertical: "middle",
        };
        worksheet.getCell("M2").alignment = {
          horizontal: "center",
          vertical: "middle",
        };
        workbook.xlsx.writeBuffer().then((data) => {
          const blob = new Blob([data], {
            type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;charset=utf-8",
          });
          saveAs(
            blob,
            "Recut Report" + " " + this.monthx + " " + this.org + "_test.xlsx"
          );
        });
        this.$q.loading.hide({});

        /*  var myChart = echarts.init(chartDom);
        var option;

        option = {
          xAxis: {
            type: "category",
            data: [
              "Jan",
              "Feb",
              "Mar",
              "Apr",
              "May",
              "Jun",
              "Jul",
              "Aug",
              "Sep",
              "Oct",
              "Nov",
              "Dec",
            ],
          },
          yAxis: {
            type: "value",
          },
          series: [
            {
              data: [
                1.2, 2.0, 1.5, 0.8, 0.7, 1.1, 1.3, 0.7, 0.9, 1.2, 1.3, 1.5,
              ],
              type: "bar",
              showBackground: true,
              backgroundStyle: {
                color: "rgba(180, 180, 180, 0.2)",
              },
            },
            {
              data: [
                1.0, 1.0, 1.0, 1.0, 1.0, 1.0, 1.0, 1.0, 1.0, 1.0, 1.0, 1.0,
              ],
              type: "line",
              color: "red",
            },
          ],
        };

        option && myChart.setOption(option);

        var myChart2 = echarts.init(chartDom2);
        var option2;

        option2 = {
          xAxis: {
            type: "category",
            data: [
              "Sewing Defect",
              "Fabric Defect",
              "Cutting Defect",
              "Heat Defect",
              "Pad Defect",
              "Lost",
              "Print Defect",
              "Fuse Defect",
              "Emb. Defect",
            ],
          },
          yAxis: {
            type: "value",
          },
          series: [
            {
              data: [0.04, 1.35, 0.21, 0.0, 0.02, 0.03, 0.16, 0.0, 0.04],
              color: "rgba(245, 178, 39, 0.8)",
              type: "bar",
              showBackground: true,
              backgroundStyle: {
                color: "rgba(180, 180, 180, 0.2)",
              },
            },
          ],
        };

        option2 && myChart2.setOption(option2); */
      }
    },
  },
  watch: {
    start(val) {
      let day = val.substring(0, 2);
      let month = val.substring(3, 5);
      let year = val.substring(6, 10);
      this.monthx = month;
      this.start_date = year + "/" + month + "/" + day;
      this.year = year;
    },
    end(val) {
      let day = val.substring(0, 2);
      let month = val.substring(3, 5);
      let year = val.substring(6, 10);
      this.end_date = year + "/" + month + "/" + day;
    },
  },
};
</script>
<style lang="sass">
.my-card
  padding-top: 10%
  padding-right: 30px
  padding-bottom: 50%
  padding-left: 80px

.center
  align:center
</style>
