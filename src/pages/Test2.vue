<template>
  <div>
    <q-btn @click="open"></q-btn>
  </div>
</template>

<script>
export default {
  methods: {
    open() {
      if (this.rowexport_use[ax].reason == "เสียจากแผนกเย็บ") {
        worksheet.getCell("T " + [ax + 4]).value = Number(
          this.rowexport_use[ax].pcs
        );
        worksheet.getCell("T " + [ax + 4]).numFmt = "0";
        worksheet.getCell("U " + [ax + 4]).value = Number(
          this.rowexport_use[ax].primary_quantity
        );
        worksheet.getCell("U " + [ax + 4]).numFmt = "0";

        if (this.rowexport_use[ax].pcs > 0) {
          this.all_result_pcs.push({
            value: this.rowexport_use[ax].pcs,
          });
        } else {
          this.all_result_pcs.push({
            value: 0,
          });
        }

        if (this.rowexport_use[ax].primary_quantity > 0) {
          this.all_result_yard.push({
            value: this.rowexport_use[ax].primary_quantity,
          });
        } else {
          this.all_result_yard.push({
            value: 0,
          });
        }
      } else if (this.rowexport_use[ax].reason2 == "เสียจากแผนกเย็บ") {
        worksheet.getCell("T " + [ax + 4]).value = Number(
          this.rowexport_use[ax].pcs2
        );
        worksheet.getCell("T " + [ax + 4]).numFmt = "0";
        worksheet.getCell("U " + [ax + 4]).value = Number(
          this.rowexport_use[ax].primary_quantity2
        );
        worksheet.getCell("U " + [ax + 4]).numFmt = "0";
        if (this.rowexport_use[ax].pcs2 > 0) {
          this.all_result_pcs.push({
            value: this.rowexport_use[ax].pcs2,
          });
        } else {
          this.all_result_pcs.push({
            value: 0,
          });
        }

        if (this.rowexport_use[ax].primary_quantity2 > 0) {
          this.all_result_yard.push({
            value: this.rowexport_use[ax].primary_quantity2,
          });
        } else {
          this.all_result_yard.push({
            value: 0,
          });
        }
      } else if (this.rowexport_use[ax].reason3 == "เสียจากแผนกเย็บ") {
        worksheet.getCell("T " + [ax + 4]).value = Number(
          this.rowexport_use[ax].pcs3
        );
        worksheet.getCell("T " + [ax + 4]).numFmt = "0";
        worksheet.getCell("U " + [ax + 4]).value = Number(
          this.rowexport_use[ax].primary_quantity3
        );
        worksheet.getCell("U " + [ax + 4]).numFmt = "0";
        if (this.rowexport_use[ax].pcs3 > 0) {
          this.all_result_pcs.push({
            value: this.rowexport_use[ax].pcs3,
          });
        } else {
          this.all_result_pcs.push({
            value: 0,
          });
        }

        if (this.rowexport_use[ax].primary_quantity3 > 0) {
          this.all_result_yard.push({
            value: this.rowexport_use[ax].primary_quantity3,
          });
        } else {
          this.all_result_yard.push({
            value: 0,
          });
        }
      } else if (this.rowexport_use[ax].reason4 == "เสียจากแผนกเย็บ") {
        worksheet.getCell("T " + [ax + 4]).value = Number(
          this.rowexport_use[ax].pcs4
        );
        worksheet.getCell("T " + [ax + 4]).numFmt = "0";
        worksheet.getCell("U " + [ax + 4]).value = Number(
          this.rowexport_use[ax].primary_quantity4
        );
        worksheet.getCell("U " + [ax + 4]).numFmt = "0";
        if (this.rowexport_use[ax].pcs4 > 0) {
          this.all_result_pcs.push({
            value: this.rowexport_use[ax].pcs4,
          });
        } else {
          this.all_result_pcs.push({
            value: 0,
          });
        }

        if (this.rowexport_use[ax].primary_quantity4 > 0) {
          this.all_result_yard.push({
            value: this.rowexport_use[ax].primary_quantity4,
          });
        } else {
          this.all_result_yard.push({
            value: 0,
          });
        }
      } else if (this.rowexport_use[ax].reason5 == "เสียจากแผนกเย็บ") {
        worksheet.getCell("T " + [ax + 4]).value = Number(
          this.rowexport_use[ax].pcs5
        );
        worksheet.getCell("T " + [ax + 4]).numFmt = "0";
        worksheet.getCell("U " + [ax + 4]).value = Number(
          this.rowexport_use[ax].primary_quantity5
        );
        worksheet.getCell("U " + [ax + 4]).numFmt = "0";
        if (this.rowexport_use[ax].pcs5 > 0) {
          this.all_result_pcs.push({
            value: this.rowexport_use[ax].pcs5,
          });
        } else {
          this.all_result_pcs.push({
            value: 0,
          });
        }

        if (this.rowexport_use[ax].primary_quantity5 > 0) {
          this.all_result_yard.push({
            value: this.rowexport_use[ax].primary_quantity5,
          });
        } else {
          this.all_result_yard.push({
            value: 0,
          });
        }
      } else if (this.rowexport_use[ax].reason7 == "เสียจากแผนกเย็บ") {
        worksheet.getCell("T " + [ax + 4]).value = Number(
          this.rowexport_use[ax].pcs7
        );
        worksheet.getCell("T " + [ax + 4]).numFmt = "0";
        worksheet.getCell("U " + [ax + 4]).value = Number(
          this.rowexport_use[ax].primary_quantity7
        );
        worksheet.getCell("U " + [ax + 4]).numFmt = "0";
        if (this.rowexport_use[ax].pcs7 > 0) {
          this.all_result_pcs.push({
            value: this.rowexport_use[ax].pcs7,
          });
        } else {
          this.all_result_pcs.push({
            value: 0,
          });
        }

        if (this.rowexport_use[ax].primary_quantity7 > 0) {
          this.all_result_yard.push({
            value: this.rowexport_use[ax].primary_quantity7,
          });
        } else {
          this.all_result_yard.push({
            value: 0,
          });
        }
      } else if (this.rowexport_use[ax].reason8 == "เสียจากแผนกเย็บ") {
        worksheet.getCell("T " + [ax + 4]).value = Number(
          this.rowexport_use[ax].pcs8
        );
        worksheet.getCell("T " + [ax + 4]).numFmt = "0";
        worksheet.getCell("U " + [ax + 4]).value = Number(
          this.rowexport_use[ax].primary_quantity8
        );
        worksheet.getCell("U " + [ax + 4]).numFmt = "0";
        if (this.rowexport_use[ax].pcs8 > 0) {
          this.all_result_pcs.push({
            value: this.rowexport_use[ax].pcs8,
          });
        } else {
          this.all_result_pcs.push({
            value: 0,
          });
        }

        if (this.rowexport_use[ax].primary_quantity8 > 0) {
          this.all_result_yard.push({
            value: this.rowexport_use[ax].primary_quantity8,
          });
        } else {
          this.all_result_yard.push({
            value: 0,
          });
        }
      } else if (this.rowexport_use[ax].reason9 == "เสียจากแผนกเย็บ") {
        worksheet.getCell("T " + [ax + 4]).value = Number(
          this.rowexport_use[ax].pcs9
        );
        worksheet.getCell("T " + [ax + 4]).numFmt = "0";
        worksheet.getCell("U " + [ax + 4]).value = Number(
          this.rowexport_use[ax].primary_quantity9
        );
        worksheet.getCell("U " + [ax + 4]).numFmt = "0";
        if (this.rowexport_use[ax].pcs9 > 0) {
          this.all_result_pcs.push({
            value: this.rowexport_use[ax].pcs9,
          });
        } else {
          this.all_result_pcs.push({
            value: 0,
          });
        }

        if (this.rowexport_use[ax].primary_quantity9 > 0) {
          this.all_result_yard.push({
            value: this.rowexport_use[ax].primary_quantity9,
          });
        } else {
          this.all_result_yard.push({
            value: 0,
          });
        }
      } else if (this.rowexport_use[ax].reason10 == "เสียจากแผนกเย็บ") {
        worksheet.getCell("T " + [ax + 4]).value = Number(
          this.rowexport_use[ax].pcs10
        );
        worksheet.getCell("T " + [ax + 4]).numFmt = "0";
        worksheet.getCell("U " + [ax + 4]).value = Number(
          this.rowexport_use[ax].primary_quantity10
        );
        worksheet.getCell("U " + [ax + 4]).numFmt = "0";
        if (this.rowexport_use[ax].pcs10 > 0) {
          this.all_result_pcs.push({
            value: this.rowexport_use[ax].pcs10,
          });
        } else {
          this.all_result_pcs.push({
            value: 0,
          });
        }

        if (this.rowexport_use[ax].primary_quantity10 > 0) {
          this.all_result_yard.push({
            value: this.rowexport_use[ax].primary_quantity10,
          });
        } else {
          this.all_result_yard.push({
            value: 0,
          });
        }
      } else {
        worksheet.getCell("T " + [ax + 4]).value = "-";
        worksheet.getCell("U " + [ax + 4]).value = "-";
        this.all_result_pcs.push({
          value: 0,
        });
        this.all_result_yard.push({
          value: 0,
        });
      }

      if (this.rowexport_use[ax].reason == "เสียจากผ้าตำหนิ") {
        worksheet.getCell("V " + [ax + 4]).value = Number(
          this.rowexport_use[ax].pcs
        );
        worksheet.getCell("V " + [ax + 4]).numFmt = "0";
        worksheet.getCell("W " + [ax + 4]).value = Number(
          this.rowexport_use[ax].primary_quantity
        );
        worksheet.getCell("W " + [ax + 4]).numFmt = "0";
        if (this.rowexport_use[ax].pcs > 0) {
          this.all_result_pcs2.push({
            value: this.rowexport_use[ax].pcs,
          });
        } else {
          this.all_result_pcs2.push({
            value: 0,
          });
        }

        if (this.rowexport_use[ax].primary_quantity > 0) {
          this.all_result_yard2.push({
            value: this.rowexport_use[ax].primary_quantity,
          });
        } else {
          this.all_result_yard2.push({
            value: 0,
          });
        }
      } else if (this.rowexport_use[ax].reason2 == "เสียจากผ้าตำหนิ") {
        worksheet.getCell("V " + [ax + 4]).value = Number(
          this.rowexport_use[ax].pcs2
        );
        worksheet.getCell("V " + [ax + 4]).numFmt = "0";
        worksheet.getCell("W " + [ax + 4]).value = Number(
          this.rowexport_use[ax].primary_quantity2
        );
        worksheet.getCell("W " + [ax + 4]).numFmt = "0";
        if (this.rowexport_use[ax].pcs2 > 0) {
          this.all_result_pcs2.push({
            value: this.rowexport_use[ax].pcs2,
          });
        } else {
          this.all_result_pcs2.push({
            value: 0,
          });
        }

        if (this.rowexport_use[ax].primary_quantity2 > 0) {
          this.all_result_yard2.push({
            value: this.rowexport_use[ax].primary_quantity2,
          });
        } else {
          this.all_result_yard2.push({
            value: 0,
          });
        }
      } else if (this.rowexport_use[ax].reason3 == "เสียจากผ้าตำหนิ") {
        worksheet.getCell("V " + [ax + 4]).value = Number(
          this.rowexport_use[ax].pcs3
        );
        worksheet.getCell("V " + [ax + 4]).numFmt = "0";
        worksheet.getCell("W " + [ax + 4]).value = Number(
          this.rowexport_use[ax].primary_quantity3
        );
        worksheet.getCell("W " + [ax + 4]).numFmt = "0";
        if (this.rowexport_use[ax].pcs3 > 0) {
          this.all_result_pcs2.push({
            value: this.rowexport_use[ax].pcs3,
          });
        } else {
          this.all_result_pcs2.push({
            value: 0,
          });
        }

        if (this.rowexport_use[ax].primary_quantity3 > 0) {
          this.all_result_yard2.push({
            value: this.rowexport_use[ax].primary_quantity3,
          });
        } else {
          this.all_result_yard2.push({
            value: 0,
          });
        }
      } else if (this.rowexport_use[ax].reason4 == "เสียจากผ้าตำหนิ") {
        worksheet.getCell("V " + [ax + 4]).value = Number(
          this.rowexport_use[ax].pcs4
        );
        worksheet.getCell("V " + [ax + 4]).numFmt = "0";
        worksheet.getCell("W " + [ax + 4]).value = Number(
          this.rowexport_use[ax].primary_quantity4
        );
        worksheet.getCell("W " + [ax + 4]).numFmt = "0";
        if (this.rowexport_use[ax].pcs4 > 0) {
          this.all_result_pcs2.push({
            value: this.rowexport_use[ax].pcs4,
          });
        } else {
          this.all_result_pcs2.push({
            value: 0,
          });
        }

        if (this.rowexport_use[ax].primary_quantity4 > 0) {
          this.all_result_yard2.push({
            value: this.rowexport_use[ax].primary_quantity4,
          });
        } else {
          this.all_result_yard2.push({
            value: 0,
          });
        }
      } else if (this.rowexport_use[ax].reason5 == "เสียจากผ้าตำหนิ") {
        worksheet.getCell("V " + [ax + 4]).value = Number(
          this.rowexport_use[ax].pcs5
        );
        worksheet.getCell("V " + [ax + 4]).numFmt = "0";
        worksheet.getCell("W " + [ax + 4]).value = Number(
          this.rowexport_use[ax].primary_quantity5
        );
        worksheet.getCell("W " + [ax + 4]).numFmt = "0";
        if (this.rowexport_use[ax].pcs5 > 0) {
          this.all_result_pcs2.push({
            value: this.rowexport_use[ax].pcs5,
          });
        } else {
          this.all_result_pcs2.push({
            value: 0,
          });
        }

        if (this.rowexport_use[ax].primary_quantity5 > 0) {
          this.all_result_yard2.push({
            value: this.rowexport_use[ax].primary_quantity5,
          });
        } else {
          this.all_result_yard2.push({
            value: 0,
          });
        }
      } else if (this.rowexport_use[ax].reason6 == "เสียจากผ้าตำหนิ") {
        worksheet.getCell("V " + [ax + 4]).value = Number(
          this.rowexport_use[ax].pcs6
        );
        worksheet.getCell("V " + [ax + 4]).numFmt = "0";
        worksheet.getCell("W " + [ax + 4]).value = Number(
          this.rowexport_use[ax].primary_quantity6
        );
        worksheet.getCell("W " + [ax + 4]).numFmt = "0";
        if (this.rowexport_use[ax].pcs6 > 0) {
          this.all_result_pcs2.push({
            value: this.rowexport_use[ax].pcs6,
          });
        } else {
          this.all_result_pcs2.push({
            value: 0,
          });
        }

        if (this.rowexport_use[ax].primary_quantity6 > 0) {
          this.all_result_yard2.push({
            value: this.rowexport_use[ax].primary_quantity6,
          });
        } else {
          this.all_result_yard2.push({
            value: 0,
          });
        }
      } else if (this.rowexport_use[ax].reason7 == "เสียจากผ้าตำหนิ") {
        worksheet.getCell("V " + [ax + 4]).value = Number(
          this.rowexport_use[ax].pcs7
        );
        worksheet.getCell("V " + [ax + 4]).numFmt = "0";
        worksheet.getCell("W " + [ax + 4]).value = Number(
          this.rowexport_use[ax].primary_quantity7
        );
        worksheet.getCell("W " + [ax + 4]).numFmt = "0";
        if (this.rowexport_use[ax].pcs7 > 0) {
          this.all_result_pcs2.push({
            value: this.rowexport_use[ax].pcs7,
          });
        } else {
          this.all_result_pcs2.push({
            value: 0,
          });
        }

        if (this.rowexport_use[ax].primary_quantity7 > 0) {
          this.all_result_yard2.push({
            value: this.rowexport_use[ax].primary_quantity7,
          });
        } else {
          this.all_result_yard2.push({
            value: 0,
          });
        }
      } else if (this.rowexport_use[ax].reason8 == "เสียจากผ้าตำหนิ") {
        worksheet.getCell("V " + [ax + 4]).value = Number(
          this.rowexport_use[ax].pcs8
        );
        worksheet.getCell("V " + [ax + 4]).numFmt = "0";
        worksheet.getCell("W " + [ax + 4]).value = Number(
          this.rowexport_use[ax].primary_quantity8
        );
        worksheet.getCell("W " + [ax + 4]).numFmt = "0";
        if (this.rowexport_use[ax].pcs8 > 0) {
          this.all_result_pcs2.push({
            value: this.rowexport_use[ax].pcs8,
          });
        } else {
          this.all_result_pcs2.push({
            value: 0,
          });
        }

        if (this.rowexport_use[ax].primary_quantity8 > 0) {
          this.all_result_yard2.push({
            value: this.rowexport_use[ax].primary_quantity8,
          });
        } else {
          this.all_result_yard2.push({
            value: 0,
          });
        }
      } else if (this.rowexport_use[ax].reason9 == "เสียจากผ้าตำหนิ") {
        worksheet.getCell("V " + [ax + 4]).value = Number(
          this.rowexport_use[ax].pcs9
        );
        worksheet.getCell("V " + [ax + 4]).numFmt = "0";
        worksheet.getCell("W " + [ax + 4]).value = Number(
          this.rowexport_use[ax].primary_quantity9
        );
        worksheet.getCell("W " + [ax + 4]).numFmt = "0";
        if (this.rowexport_use[ax].pcs9 > 0) {
          this.all_result_pcs2.push({
            value: this.rowexport_use[ax].pcs9,
          });
        } else {
          this.all_result_pcs2.push({
            value: 0,
          });
        }

        if (this.rowexport_use[ax].primary_quantity9 > 0) {
          this.all_result_yard2.push({
            value: this.rowexport_use[ax].primary_quantity9,
          });
        } else {
          this.all_result_yard2.push({
            value: 0,
          });
        }
      } else if (this.rowexport_use[ax].reason10 == "เสียจากผ้าตำหนิ") {
        worksheet.getCell("V " + [ax + 4]).value = Number(
          this.rowexport_use[ax].pcs10
        );
        worksheet.getCell("V " + [ax + 4]).numFmt = "0";
        worksheet.getCell("W " + [ax + 4]).value = Number(
          this.rowexport_use[ax].primary_quantity10
        );
        worksheet.getCell("W " + [ax + 4]).numFmt = "0";
        if (this.rowexport_use[ax].pcs10 > 0) {
          this.all_result_pcs2.push({
            value: this.rowexport_use[ax].pcs10,
          });
        } else {
          this.all_result_pcs2.push({
            value: 0,
          });
        }

        if (this.rowexport_use[ax].primary_quantity10 > 0) {
          this.all_result_yard2.push({
            value: this.rowexport_use[ax].primary_quantity10,
          });
        } else {
          this.all_result_yard2.push({
            value: 0,
          });
        }
      } else {
        worksheet.getCell("V " + [ax + 4]).value = "-";
        worksheet.getCell("W " + [ax + 4]).value = "-";
        this.all_result_pcs2.push({
          value: 0,
        });
        this.all_result_yard2.push({
          value: 0,
        });
      }

      if (this.rowexport_use[ax].reason == "งานตัดคัดชิ้น") {
        worksheet.getCell("X " + [ax + 4]).value = Number(
          this.rowexport_use[ax].pcs
        );
        worksheet.getCell("X " + [ax + 4]).numFmt = "0";
        worksheet.getCell("Y " + [ax + 4]).value = Number(
          this.rowexport_use[ax].primary_quantity
        );
        worksheet.getCell("Y " + [ax + 4]).numFmt = "0";
        if (this.rowexport_use[ax].pcs > 0) {
          this.all_result_pcs3.push({
            value: this.rowexport_use[ax].pcs,
          });
        } else {
          this.all_result_pcs3.push({
            value: 0,
          });
        }

        if (this.rowexport_use[ax].primary_quantity > 0) {
          this.all_result_yard3.push({
            value: this.rowexport_use[ax].primary_quantity,
          });
        } else {
          this.all_result_yard3.push({
            value: 0,
          });
        }
      } else if (this.rowexport_use[ax].reason2 == "งานตัดคัดชิ้น") {
        worksheet.getCell("X " + [ax + 4]).value = Number(
          this.rowexport_use[ax].pcs2
        );
        worksheet.getCell("X " + [ax + 4]).numFmt = "0";
        worksheet.getCell("Y " + [ax + 4]).value = Number(
          this.rowexport_use[ax].primary_quantity2
        );
        worksheet.getCell("Y " + [ax + 4]).numFmt = "0";
        if (this.rowexport_use[ax].pcs > 0) {
          this.all_result_pcs3.push({
            value: this.rowexport_use[ax].pcs2,
          });
        } else {
          this.all_result_pcs3.push({
            value: 0,
          });
        }

        if (this.rowexport_use[ax].primary_quantity2 > 0) {
          this.all_result_yard3.push({
            value: this.rowexport_use[ax].primary_quantity2,
          });
        } else {
          this.all_result_yard3.push({
            value: 0,
          });
        }
      } else if (this.rowexport_use[ax].reason3 == "งานตัดคัดชิ้น") {
        worksheet.getCell("X " + [ax + 4]).value = Number(
          this.rowexport_use[ax].pcs3
        );
        worksheet.getCell("X " + [ax + 4]).numFmt = "0";
        worksheet.getCell("Y " + [ax + 4]).value = Number(
          this.rowexport_use[ax].primary_quantity3
        );
        worksheet.getCell("Y " + [ax + 4]).numFmt = "0";
        if (this.rowexport_use[ax].pcs3 > 0) {
          this.all_result_pcs3.push({
            value: this.rowexport_use[ax].pcs3,
          });
        } else {
          this.all_result_pcs3.push({
            value: 0,
          });
        }

        if (this.rowexport_use[ax].primary_quantity3 > 0) {
          this.all_result_yard3.push({
            value: this.rowexport_use[ax].primary_quantity3,
          });
        } else {
          this.all_result_yard3.push({
            value: 0,
          });
        }
      } else if (this.rowexport_use[ax].reason4 == "งานตัดคัดชิ้น") {
        worksheet.getCell("X " + [ax + 4]).value = Number(
          this.rowexport_use[ax].pcs4
        );
        worksheet.getCell("X " + [ax + 4]).numFmt = "0";
        worksheet.getCell("Y " + [ax + 4]).value = Number(
          this.rowexport_use[ax].primary_quantity4
        );
        worksheet.getCell("Y " + [ax + 4]).numFmt = "0";
        if (this.rowexport_use[ax].pcs4 > 0) {
          this.all_result_pcs3.push({
            value: this.rowexport_use[ax].pcs4,
          });
        } else {
          this.all_result_pcs3.push({
            value: 0,
          });
        }

        if (this.rowexport_use[ax].primary_quantity4 > 0) {
          this.all_result_yard3.push({
            value: this.rowexport_use[ax].primary_quantity4,
          });
        } else {
          this.all_result_yard3.push({
            value: 0,
          });
        }
      } else if (this.rowexport_use[ax].reason5 == "งานตัดคัดชิ้น") {
        worksheet.getCell("X " + [ax + 4]).value = Number(
          this.rowexport_use[ax].pcs5
        );
        worksheet.getCell("X " + [ax + 4]).numFmt = "0";
        worksheet.getCell("Y " + [ax + 4]).value = Number(
          this.rowexport_use[ax].primary_quantity5
        );
        worksheet.getCell("Y " + [ax + 4]).numFmt = "0";
        if (this.rowexport_use[ax].pcs5 > 0) {
          this.all_result_pcs3.push({
            value: this.rowexport_use[ax].pcs5,
          });
        } else {
          this.all_result_pcs3.push({
            value: 0,
          });
        }

        if (this.rowexport_use[ax].primary_quantity5 > 0) {
          this.all_result_yard3.push({
            value: this.rowexport_use[ax].primary_quantity5,
          });
        } else {
          this.all_result_yard3.push({
            value: 0,
          });
        }
      } else if (this.rowexport_use[ax].reason6 == "งานตัดคัดชิ้น") {
        worksheet.getCell("X " + [ax + 4]).value = Number(
          this.rowexport_use[ax].pcs6
        );
        worksheet.getCell("X " + [ax + 4]).numFmt = "0";
        worksheet.getCell("Y " + [ax + 4]).value = Number(
          this.rowexport_use[ax].primary_quantity6
        );
        worksheet.getCell("Y " + [ax + 4]).numFmt = "0";
        if (this.rowexport_use[ax].pcs6 > 0) {
          this.all_result_pcs3.push({
            value: this.rowexport_use[ax].pcs6,
          });
        } else {
          this.all_result_pcs3.push({
            value: 0,
          });
        }

        if (this.rowexport_use[ax].primary_quantity6 > 0) {
          this.all_result_yard3.push({
            value: this.rowexport_use[ax].primary_quantity6,
          });
        } else {
          this.all_result_yard3.push({
            value: 0,
          });
        }
      } else if (this.rowexport_use[ax].reason7 == "งานตัดคัดชิ้น") {
        worksheet.getCell("X " + [ax + 4]).value = Number(
          this.rowexport_use[ax].pcs7
        );
        worksheet.getCell("X " + [ax + 4]).numFmt = "0";
        worksheet.getCell("Y " + [ax + 4]).value = Number(
          this.rowexport_use[ax].primary_quantity7
        );
        worksheet.getCell("Y " + [ax + 4]).numFmt = "0";
        if (this.rowexport_use[ax].pcs7 > 0) {
          this.all_result_pcs3.push({
            value: this.rowexport_use[ax].pcs7,
          });
        } else {
          this.all_result_pcs3.push({
            value: 0,
          });
        }

        if (this.rowexport_use[ax].primary_quantity7 > 0) {
          this.all_result_yard3.push({
            value: this.rowexport_use[ax].primary_quantity7,
          });
        } else {
          this.all_result_yard3.push({
            value: 0,
          });
        }
      } else if (this.rowexport_use[ax].reason8 == "งานตัดคัดชิ้น") {
        worksheet.getCell("X " + [ax + 4]).value = Number(
          this.rowexport_use[ax].pcs8
        );
        worksheet.getCell("X " + [ax + 4]).numFmt = "0";
        worksheet.getCell("Y " + [ax + 4]).value = Number(
          this.rowexport_use[ax].primary_quantity8
        );
        worksheet.getCell("Y " + [ax + 4]).numFmt = "0";
        if (this.rowexport_use[ax].pcs8 > 0) {
          this.all_result_pcs3.push({
            value: this.rowexport_use[ax].pcs8,
          });
        } else {
          this.all_result_pcs3.push({
            value: 0,
          });
        }

        if (this.rowexport_use[ax].primary_quantity8 > 0) {
          this.all_result_yard3.push({
            value: this.rowexport_use[ax].primary_quantity8,
          });
        } else {
          this.all_result_yard3.push({
            value: 0,
          });
        }
      } else if (this.rowexport_use[ax].reason9 == "งานตัดคัดชิ้น") {
        worksheet.getCell("X " + [ax + 4]).value = Number(
          this.rowexport_use[ax].pcs9
        );
        worksheet.getCell("X " + [ax + 4]).numFmt = "0";
        worksheet.getCell("Y " + [ax + 4]).value = Number(
          this.rowexport_use[ax].primary_quantity9
        );
        worksheet.getCell("Y " + [ax + 4]).numFmt = "0";
        if (this.rowexport_use[ax].pcs9 > 0) {
          this.all_result_pcs3.push({
            value: this.rowexport_use[ax].pcs9,
          });
        } else {
          this.all_result_pcs3.push({
            value: 0,
          });
        }

        if (this.rowexport_use[ax].primary_quantity9 > 0) {
          this.all_result_yard3.push({
            value: this.rowexport_use[ax].primary_quantity9,
          });
        } else {
          this.all_result_yard3.push({
            value: 0,
          });
        }
      } else if (this.rowexport_use[ax].reason10 == "งานตัดคัดชิ้น") {
        worksheet.getCell("X " + [ax + 4]).value = Number(
          this.rowexport_use[ax].pcs10
        );
        worksheet.getCell("X " + [ax + 4]).numFmt = "0";
        worksheet.getCell("Y " + [ax + 4]).value = Number(
          this.rowexport_use[ax].primary_quantity10
        );
        worksheet.getCell("Y " + [ax + 4]).numFmt = "0";
        if (this.rowexport_use[ax].pcs10 > 0) {
          this.all_result_pcs3.push({
            value: this.rowexport_use[ax].pcs10,
          });
        } else {
          this.all_result_pcs3.push({
            value: 0,
          });
        }

        if (this.rowexport_use[ax].primary_quantity10 > 0) {
          this.all_result_yard3.push({
            value: this.rowexport_use[ax].primary_quantity10,
          });
        } else {
          this.all_result_yard3.push({
            value: 0,
          });
        }
      } else {
        worksheet.getCell("X " + [ax + 4]).value = "-";
        worksheet.getCell("Y " + [ax + 4]).value = "-";
        this.all_result_pcs3.push({
          value: 0,
        });
        this.all_result_yard3.push({
          value: 0,
        });
      }

      if (this.rowexport_use[ax].reason == "เสียจากแผนกตัด") {
        worksheet.getCell("Z " + [ax + 4]).value = Number(
          this.rowexport_use[ax].pcs
        );
        worksheet.getCell("Z " + [ax + 4]).numFmt = "0";
        worksheet.getCell("AA " + [ax + 4]).value = Number(
          this.rowexport_use[ax].primary_quantity
        );
        worksheet.getCell("AA " + [ax + 4]).numFmt = "0";
        if (this.rowexport_use[ax].pcs > 0) {
          this.all_result_pcs4.push({
            value: this.rowexport_use[ax].pcs,
          });
        } else {
          this.all_result_pcs4.push({
            value: 0,
          });
        }

        if (this.rowexport_use[ax].primary_quantity > 0) {
          this.all_result_yard4.push({
            value: this.rowexport_use[ax].primary_quantity,
          });
        } else {
          this.all_result_yard4.push({
            value: 0,
          });
        }
      } else if (this.rowexport_use[ax].reason2 == "เสียจากแผนกตัด") {
        worksheet.getCell("Z " + [ax + 4]).value = Number(
          this.rowexport_use[ax].pcs2
        );
        worksheet.getCell("Z " + [ax + 4]).numFmt = "0";
        worksheet.getCell("AA " + [ax + 4]).value = Number(
          this.rowexport_use[ax].primary_quantity2
        );
        worksheet.getCell("AA " + [ax + 4]).numFmt = "0";
        if (this.rowexport_use[ax].pcs2 > 0) {
          this.all_result_pcs4.push({
            value: this.rowexport_use[ax].pcs2,
          });
        } else {
          this.all_result_pcs4.push({
            value: 0,
          });
        }

        if (this.rowexport_use[ax].primary_quantity2 > 0) {
          this.all_result_yard4.push({
            value: this.rowexport_use[ax].primary_quantity2,
          });
        } else {
          this.all_result_yard4.push({
            value: 0,
          });
        }
      } else if (this.rowexport_use[ax].reason3 == "เสียจากแผนกตัด") {
        worksheet.getCell("Z " + [ax + 4]).value = Number(
          this.rowexport_use[ax].pcs3
        );
        worksheet.getCell("Z " + [ax + 4]).numFmt = "0";
        worksheet.getCell("AA " + [ax + 4]).value = Number(
          this.rowexport_use[ax].primary_quantity3
        );
        worksheet.getCell("AA " + [ax + 4]).numFmt = "0";
        if (this.rowexport_use[ax].pcs3 > 0) {
          this.all_result_pcs4.push({
            value: this.rowexport_use[ax].pcs3,
          });
        } else {
          this.all_result_pcs4.push({
            value: 0,
          });
        }

        if (this.rowexport_use[ax].primary_quantity3 > 0) {
          this.all_result_yard4.push({
            value: this.rowexport_use[ax].primary_quantity3,
          });
        } else {
          this.all_result_yard4.push({
            value: 0,
          });
        }
      } else if (this.rowexport_use[ax].reason4 == "เสียจากแผนกตัด") {
        worksheet.getCell("Z " + [ax + 4]).value = Number(
          this.rowexport_use[ax].pcs4
        );
        worksheet.getCell("Z " + [ax + 4]).numFmt = "0";
        worksheet.getCell("AA " + [ax + 4]).value = Number(
          this.rowexport_use[ax].primary_quantity4
        );
        worksheet.getCell("AA " + [ax + 4]).numFmt = "0";
        if (this.rowexport_use[ax].pcs4 > 0) {
          this.all_result_pcs4.push({
            value: this.rowexport_use[ax].pcs4,
          });
        } else {
          this.all_result_pcs4.push({
            value: 0,
          });
        }

        if (this.rowexport_use[ax].primary_quantity4 > 0) {
          this.all_result_yard4.push({
            value: this.rowexport_use[ax].primary_quantity4,
          });
        } else {
          this.all_result_yard4.push({
            value: 0,
          });
        }
      } else if (this.rowexport_use[ax].reason5 == "เสียจากแผนกตัด") {
        worksheet.getCell("Z " + [ax + 4]).value = Number(
          this.rowexport_use[ax].pcs5
        );
        worksheet.getCell("Z " + [ax + 4]).numFmt = "0";
        worksheet.getCell("AA " + [ax + 4]).value = Number(
          this.rowexport_use[ax].primary_quantity5
        );
        worksheet.getCell("AA " + [ax + 4]).numFmt = "0";
        if (this.rowexport_use[ax].pcs5 > 0) {
          this.all_result_pcs4.push({
            value: this.rowexport_use[ax].pcs5,
          });
        } else {
          this.all_result_pcs4.push({
            value: 0,
          });
        }

        if (this.rowexport_use[ax].primary_quantity5 > 0) {
          this.all_result_yard4.push({
            value: this.rowexport_use[ax].primary_quantity5,
          });
        } else {
          this.all_result_yard4.push({
            value: 0,
          });
        }
      } else if (this.rowexport_use[ax].reason6 == "เสียจากแผนกตัด") {
        worksheet.getCell("Z " + [ax + 4]).value = Number(
          this.rowexport_use[ax].pcs6
        );
        worksheet.getCell("Z " + [ax + 4]).numFmt = "0";
        worksheet.getCell("AA " + [ax + 4]).value = Number(
          this.rowexport_use[ax].primary_quantity6
        );
        worksheet.getCell("AA " + [ax + 4]).numFmt = "0";
        if (this.rowexport_use[ax].pcs6 > 0) {
          this.all_result_pcs4.push({
            value: this.rowexport_use[ax].pcs6,
          });
        } else {
          this.all_result_pcs4.push({
            value: 0,
          });
        }

        if (this.rowexport_use[ax].primary_quantity6 > 0) {
          this.all_result_yard4.push({
            value: this.rowexport_use[ax].primary_quantity6,
          });
        } else {
          this.all_result_yard4.push({
            value: 0,
          });
        }
      } else if (this.rowexport_use[ax].reason7 == "เสียจากแผนกตัด") {
        worksheet.getCell("Z " + [ax + 4]).value = Number(
          this.rowexport_use[ax].pcs7
        );
        worksheet.getCell("Z " + [ax + 4]).numFmt = "0";
        worksheet.getCell("AA " + [ax + 4]).value = Number(
          this.rowexport_use[ax].primary_quantity7
        );
        worksheet.getCell("AA " + [ax + 4]).numFmt = "0";
        if (this.rowexport_use[ax].pcs7 > 0) {
          this.all_result_pcs4.push({
            value: this.rowexport_use[ax].pcs7,
          });
        } else {
          this.all_result_pcs4.push({
            value: 0,
          });
        }

        if (this.rowexport_use[ax].primary_quantity7 > 0) {
          this.all_result_yard4.push({
            value: this.rowexport_use[ax].primary_quantity7,
          });
        } else {
          this.all_result_yard4.push({
            value: 0,
          });
        }
      } else if (this.rowexport_use[ax].reason8 == "เสียจากแผนกตัด") {
        worksheet.getCell("Z " + [ax + 4]).value = Number(
          this.rowexport_use[ax].pcs8
        );
        worksheet.getCell("Z " + [ax + 4]).numFmt = "0";
        worksheet.getCell("AA " + [ax + 4]).value = Number(
          this.rowexport_use[ax].primary_quantity8
        );
        worksheet.getCell("AA " + [ax + 4]).numFmt = "0";
        if (this.rowexport_use[ax].pcs8 > 0) {
          this.all_result_pcs4.push({
            value: this.rowexport_use[ax].pcs8,
          });
        } else {
          this.all_result_pcs4.push({
            value: 0,
          });
        }

        if (this.rowexport_use[ax].primary_quantity8 > 0) {
          this.all_result_yard4.push({
            value: this.rowexport_use[ax].primary_quantity8,
          });
        } else {
          this.all_result_yard4.push({
            value: 0,
          });
        }
      } else if (this.rowexport_use[ax].reason9 == "เสียจากแผนกตัด") {
        worksheet.getCell("Z " + [ax + 4]).value = Number(
          this.rowexport_use[ax].pcs9
        );
        worksheet.getCell("Z " + [ax + 4]).numFmt = "0";
        worksheet.getCell("AA " + [ax + 4]).value = Number(
          this.rowexport_use[ax].primary_quantity9
        );
        worksheet.getCell("AA " + [ax + 4]).numFmt = "0";
        if (this.rowexport_use[ax].pcs9 > 0) {
          this.all_result_pcs4.push({
            value: this.rowexport_use[ax].pcs9,
          });
        } else {
          this.all_result_pcs4.push({
            value: 0,
          });
        }

        if (this.rowexport_use[ax].primary_quantity9 > 0) {
          this.all_result_yard4.push({
            value: this.rowexport_use[ax].primary_quantity9,
          });
        } else {
          this.all_result_yard4.push({
            value: 0,
          });
        }
      } else if (this.rowexport_use[ax].reason10 == "เสียจากแผนกตัด") {
        worksheet.getCell("Z " + [ax + 4]).value = Number(
          this.rowexport_use[ax].pcs10
        );
        worksheet.getCell("Z " + [ax + 4]).numFmt = "0";
        worksheet.getCell("AA " + [ax + 4]).value = Number(
          this.rowexport_use[ax].primary_quantity10
        );
        worksheet.getCell("AA " + [ax + 4]).numFmt = "0";
        if (this.rowexport_use[ax].pcs10 > 0) {
          this.all_result_pcs4.push({
            value: this.rowexport_use[ax].pcs10,
          });
        } else {
          this.all_result_pcs4.push({
            value: 0,
          });
        }

        if (this.rowexport_use[ax].primary_quantity10 > 0) {
          this.all_result_yard4.push({
            value: this.rowexport_use[ax].primary_quantity10,
          });
        } else {
          this.all_result_yard4.push({
            value: 0,
          });
        }
      } else {
        worksheet.getCell("Z " + [ax + 4]).value = "-";
        worksheet.getCell("AA " + [ax + 4]).value = "-";
        this.all_result_pcs4.push({
          value: 0,
        });
        this.all_result_yard4.push({
          value: 0,
        });
      }

      if (this.rowexport_use[ax].reason == "เสียจากแผนกรีด") {
        worksheet.getCell("AB " + [ax + 4]).value = Number(
          this.rowexport_use[ax].pcs
        );
        worksheet.getCell("AB " + [ax + 4]).numFmt = "0";
        worksheet.getCell("AC " + [ax + 4]).value = Number(
          this.rowexport_use[ax].primary_quantity
        );
        worksheet.getCell("AC " + [ax + 4]).numFmt = "0";
        if (this.rowexport_use[ax].pcs > 0) {
          this.all_result_pcs5.push({
            value: this.rowexport_use[ax].pcs,
          });
        } else {
          this.all_result_pcs5.push({
            value: 0,
          });
        }

        if (this.rowexport_use[ax].primary_quantity > 0) {
          this.all_result_yard5.push({
            value: this.rowexport_use[ax].primary_quantity,
          });
        } else {
          this.all_result_yard5.push({
            value: 0,
          });
        }
      } else if (this.rowexport_use[ax].reason2 == "เสียจากแผนกรีด") {
        worksheet.getCell("AB " + [ax + 4]).value = Number(
          this.rowexport_use[ax].pcs2
        );
        worksheet.getCell("AB " + [ax + 4]).numFmt = "0";
        worksheet.getCell("AC " + [ax + 4]).value = Number(
          this.rowexport_use[ax].primary_quantity2
        );
        worksheet.getCell("AC " + [ax + 4]).numFmt = "0";
        if (this.rowexport_use[ax].pcs2 > 0) {
          this.all_result_pcs5.push({
            value: this.rowexport_use[ax].pcs2,
          });
        } else {
          this.all_result_pcs5.push({
            value: 0,
          });
        }

        if (this.rowexport_use[ax].primary_quantity2 > 0) {
          this.all_result_yard5.push({
            value: this.rowexport_use[ax].primary_quantity2,
          });
        } else {
          this.all_result_yard5.push({
            value: 0,
          });
        }
      } else if (this.rowexport_use[ax].reason3 == "เสียจากแผนกรีด") {
        worksheet.getCell("AB " + [ax + 4]).value = Number(
          this.rowexport_use[ax].pcs3
        );
        worksheet.getCell("AB " + [ax + 4]).numFmt = "0";
        worksheet.getCell("AC " + [ax + 4]).value = Number(
          this.rowexport_use[ax].primary_quantity3
        );
        worksheet.getCell("AC " + [ax + 4]).numFmt = "0";
        if (this.rowexport_use[ax].pcs3 > 0) {
          this.all_result_pcs5.push({
            value: this.rowexport_use[ax].pcs3,
          });
        } else {
          this.all_result_pcs5.push({
            value: 0,
          });
        }

        if (this.rowexport_use[ax].primary_quantity3 > 0) {
          this.all_result_yard5.push({
            value: this.rowexport_use[ax].primary_quantity3,
          });
        } else {
          this.all_result_yard5.push({
            value: 0,
          });
        }
      } else if (this.rowexport_use[ax].reason4 == "เสียจากแผนกรีด") {
        worksheet.getCell("AB " + [ax + 4]).value = Number(
          this.rowexport_use[ax].pcs4
        );
        worksheet.getCell("AB " + [ax + 4]).numFmt = "0";
        worksheet.getCell("AC " + [ax + 4]).value = Number(
          this.rowexport_use[ax].primary_quantity4
        );
        worksheet.getCell("AC " + [ax + 4]).numFmt = "0";
        if (this.rowexport_use[ax].pcs4 > 0) {
          this.all_result_pcs5.push({
            value: this.rowexport_use[ax].pcs4,
          });
        } else {
          this.all_result_pcs5.push({
            value: 0,
          });
        }

        if (this.rowexport_use[ax].primary_quantity4 > 0) {
          this.all_result_yard5.push({
            value: this.rowexport_use[ax].primary_quantity4,
          });
        } else {
          this.all_result_yard5.push({
            value: 0,
          });
        }
      } else if (this.rowexport_use[ax].reason5 == "เสียจากแผนกรีด") {
        worksheet.getCell("AB " + [ax + 4]).value = Number(
          this.rowexport_use[ax].pcs5
        );
        worksheet.getCell("AB " + [ax + 4]).numFmt = "0";
        worksheet.getCell("AC " + [ax + 4]).value = Number(
          this.rowexport_use[ax].primary_quantity5
        );
        worksheet.getCell("AC " + [ax + 4]).numFmt = "0";
        if (this.rowexport_use[ax].pcs5 > 0) {
          this.all_result_pcs5.push({
            value: this.rowexport_use[ax].pcs5,
          });
        } else {
          this.all_result_pcs5.push({
            value: 0,
          });
        }

        if (this.rowexport_use[ax].primary_quantity5 > 0) {
          this.all_result_yard5.push({
            value: this.rowexport_use[ax].primary_quantity5,
          });
        } else {
          this.all_result_yard5.push({
            value: 0,
          });
        }
      } else if (this.rowexport_use[ax].reason6 == "เสียจากแผนกรีด") {
        worksheet.getCell("AB " + [ax + 4]).value = Number(
          this.rowexport_use[ax].pcs6
        );
        worksheet.getCell("AB " + [ax + 4]).numFmt = "0";
        worksheet.getCell("AC " + [ax + 4]).value = Number(
          this.rowexport_use[ax].primary_quantity6
        );
        worksheet.getCell("AC " + [ax + 4]).numFmt = "0";
        if (this.rowexport_use[ax].pcs6 > 0) {
          this.all_result_pcs5.push({
            value: this.rowexport_use[ax].pcs6,
          });
        } else {
          this.all_result_pcs5.push({
            value: 0,
          });
        }

        if (this.rowexport_use[ax].primary_quantity6 > 0) {
          this.all_result_yard5.push({
            value: this.rowexport_use[ax].primary_quantity6,
          });
        } else {
          this.all_result_yard5.push({
            value: 0,
          });
        }
      } else if (this.rowexport_use[ax].reason7 == "เสียจากแผนกรีด") {
        worksheet.getCell("AB " + [ax + 4]).value = Number(
          this.rowexport_use[ax].pcs7
        );
        worksheet.getCell("AB " + [ax + 4]).numFmt = "0";
        worksheet.getCell("AC " + [ax + 4]).value = Number(
          this.rowexport_use[ax].primary_quantity7
        );
        worksheet.getCell("AC " + [ax + 4]).numFmt = "0";
        if (this.rowexport_use[ax].pcs7 > 0) {
          this.all_result_pcs5.push({
            value: this.rowexport_use[ax].pcs7,
          });
        } else {
          this.all_result_pcs5.push({
            value: 0,
          });
        }

        if (this.rowexport_use[ax].primary_quantity7 > 0) {
          this.all_result_yard5.push({
            value: this.rowexport_use[ax].primary_quantity7,
          });
        } else {
          this.all_result_yard5.push({
            value: 0,
          });
        }
      } else if (this.rowexport_use[ax].reason8 == "เสียจากแผนกรีด") {
        worksheet.getCell("AB " + [ax + 4]).value = Number(
          this.rowexport_use[ax].pcs8
        );
        worksheet.getCell("AB " + [ax + 4]).numFmt = "0";
        worksheet.getCell("AC " + [ax + 4]).value = Number(
          this.rowexport_use[ax].primary_quantity8
        );
        worksheet.getCell("AC " + [ax + 4]).numFmt = "0";
        if (this.rowexport_use[ax].pcs8 > 0) {
          this.all_result_pcs5.push({
            value: this.rowexport_use[ax].pcs8,
          });
        } else {
          this.all_result_pcs5.push({
            value: 0,
          });
        }

        if (this.rowexport_use[ax].primary_quantity8 > 0) {
          this.all_result_yard5.push({
            value: this.rowexport_use[ax].primary_quantity8,
          });
        } else {
          this.all_result_yard5.push({
            value: 0,
          });
        }
      } else if (this.rowexport_use[ax].reason9 == "เสียจากแผนกรีด") {
        worksheet.getCell("AB " + [ax + 4]).value = Number(
          this.rowexport_use[ax].pcs9
        );
        worksheet.getCell("AB " + [ax + 4]).numFmt = "0";
        worksheet.getCell("AC " + [ax + 4]).value = Number(
          this.rowexport_use[ax].primary_quantity9
        );
        worksheet.getCell("AC " + [ax + 4]).numFmt = "0";
        if (this.rowexport_use[ax].pcs9 > 0) {
          this.all_result_pcs5.push({
            value: this.rowexport_use[ax].pcs9,
          });
        } else {
          this.all_result_pcs5.push({
            value: 0,
          });
        }

        if (this.rowexport_use[ax].primary_quantity9 > 0) {
          this.all_result_yard5.push({
            value: this.rowexport_use[ax].primary_quantity9,
          });
        } else {
          this.all_result_yard5.push({
            value: 0,
          });
        }
      } else if (this.rowexport_use[ax].reason10 == "เสียจากแผนกรีด") {
        worksheet.getCell("AB " + [ax + 4]).value = Number(
          this.rowexport_use[ax].pcs10
        );
        worksheet.getCell("AB " + [ax + 4]).numFmt = "0";
        worksheet.getCell("AC " + [ax + 4]).value = Number(
          this.rowexport_use[ax].primary_quantity10
        );
        worksheet.getCell("AC " + [ax + 4]).numFmt = "0";
        if (this.rowexport_use[ax].pcs10 > 0) {
          this.all_result_pcs5.push({
            value: this.rowexport_use[ax].pcs10,
          });
        } else {
          this.all_result_pcs5.push({
            value: 0,
          });
        }

        if (this.rowexport_use[ax].primary_quantity10 > 0) {
          this.all_result_yard5.push({
            value: this.rowexport_use[ax].primary_quantity10,
          });
        } else {
          this.all_result_yard5.push({
            value: 0,
          });
        }
      } else {
        worksheet.getCell("AB " + [ax + 4]).value = "-";
        worksheet.getCell("AC " + [ax + 4]).value = "-";
        this.all_result_pcs5.push({
          value: 0,
        });
        this.all_result_yard5.push({
          value: 0,
        });
      }

      if (this.rowexport_use[ax].reason == "เสียจากแผนกแพด") {
        worksheet.getCell("AD " + [ax + 4]).value = Number(
          this.rowexport_use[ax].pcs
        );
        worksheet.getCell("AD " + [ax + 4]).numFmt = "0";
        worksheet.getCell("AE " + [ax + 4]).value = Number(
          this.rowexport_use[ax].primary_quantity
        );
        worksheet.getCell("AE " + [ax + 4]).numFmt = "0";
        if (this.rowexport_use[ax].pcs > 0) {
          this.all_result_pcs6.push({
            value: this.rowexport_use[ax].pcs,
          });
        } else {
          this.all_result_pcs6.push({
            value: 0,
          });
        }

        if (this.rowexport_use[ax].primary_quantity > 0) {
          this.all_result_yard6.push({
            value: this.rowexport_use[ax].primary_quantity,
          });
        } else {
          this.all_result_yard6.push({
            value: 0,
          });
        }
      } else if (this.rowexport_use[ax].reason2 == "เสียจากแผนกแพด") {
        worksheet.getCell("AD " + [ax + 4]).value = Number(
          this.rowexport_use[ax].pcs2
        );
        worksheet.getCell("AD " + [ax + 4]).numFmt = "0";
        worksheet.getCell("AE " + [ax + 4]).value = Number(
          this.rowexport_use[ax].primary_quantity2
        );
        worksheet.getCell("AE " + [ax + 4]).numFmt = "0";
        if (this.rowexport_use[ax].pcs2 > 0) {
          this.all_result_pcs6.push({
            value: this.rowexport_use[ax].pcs2,
          });
        } else {
          this.all_result_pcs6.push({
            value: 0,
          });
        }

        if (this.rowexport_use[ax].primary_quantity2 > 0) {
          this.all_result_yard6.push({
            value: this.rowexport_use[ax].primary_quantity2,
          });
        } else {
          this.all_result_yard6.push({
            value: 0,
          });
        }
      } else if (this.rowexport_use[ax].reason3 == "เสียจากแผนกแพด") {
        worksheet.getCell("AD " + [ax + 4]).value = Number(
          this.rowexport_use[ax].pcs3
        );
        worksheet.getCell("AD " + [ax + 4]).numFmt = "0";
        worksheet.getCell("AE " + [ax + 4]).value = Number(
          this.rowexport_use[ax].primary_quantity3
        );
        worksheet.getCell("AE " + [ax + 4]).numFmt = "0";
        if (this.rowexport_use[ax].pcs3 > 0) {
          this.all_result_pcs6.push({
            value: this.rowexport_use[ax].pcs3,
          });
        } else {
          this.all_result_pcs6.push({
            value: 0,
          });
        }

        if (this.rowexport_use[ax].primary_quantity3 > 0) {
          this.all_result_yard6.push({
            value: this.rowexport_use[ax].primary_quantity3,
          });
        } else {
          this.all_result_yard6.push({
            value: 0,
          });
        }
      } else if (this.rowexport_use[ax].reason4 == "เสียจากแผนกแพด") {
        worksheet.getCell("AD " + [ax + 4]).value = Number(
          this.rowexport_use[ax].pcs4
        );
        worksheet.getCell("AD " + [ax + 4]).numFmt = "0";
        worksheet.getCell("AE " + [ax + 4]).value = Number(
          this.rowexport_use[ax].primary_quantity4
        );
        worksheet.getCell("AE " + [ax + 4]).numFmt = "0";
        if (this.rowexport_use[ax].pcs4 > 0) {
          this.all_result_pcs6.push({
            value: this.rowexport_use[ax].pcs4,
          });
        } else {
          this.all_result_pcs6.push({
            value: 0,
          });
        }

        if (this.rowexport_use[ax].primary_quantity > 0) {
          this.all_result_yard6.push({
            value: this.rowexport_use[ax].primary_quantity4,
          });
        } else {
          this.all_result_yard6.push({
            value: 0,
          });
        }
      } else if (this.rowexport_use[ax].reason5 == "เสียจากแผนกแพด") {
        worksheet.getCell("AD " + [ax + 4]).value = Number(
          this.rowexport_use[ax].pcs5
        );
        worksheet.getCell("AD " + [ax + 4]).numFmt = "0";
        worksheet.getCell("AE " + [ax + 4]).value = Number(
          this.rowexport_use[ax].primary_quantity5
        );
        worksheet.getCell("AE " + [ax + 4]).numFmt = "0";
        if (this.rowexport_use[ax].pcs5 > 0) {
          this.all_result_pcs6.push({
            value: this.rowexport_use[ax].pcs5,
          });
        } else {
          this.all_result_pcs6.push({
            value: 0,
          });
        }

        if (this.rowexport_use[ax].primary_quantity5 > 0) {
          this.all_result_yard6.push({
            value: this.rowexport_use[ax].primary_quantity5,
          });
        } else {
          this.all_result_yard6.push({
            value: 0,
          });
        }
      } else if (this.rowexport_use[ax].reason6 == "เสียจากแผนกแพด") {
        worksheet.getCell("AD " + [ax + 4]).value = Number(
          this.rowexport_use[ax].pcs6
        );
        worksheet.getCell("AD " + [ax + 4]).numFmt = "0";
        worksheet.getCell("AE " + [ax + 4]).value = Number(
          this.rowexport_use[ax].primary_quantity6
        );
        worksheet.getCell("AE " + [ax + 4]).numFmt = "0";
        if (this.rowexport_use[ax].pcs6 > 0) {
          this.all_result_pcs6.push({
            value: this.rowexport_use[ax].pcs6,
          });
        } else {
          this.all_result_pcs6.push({
            value: 0,
          });
        }

        if (this.rowexport_use[ax].primary_quantity6 > 0) {
          this.all_result_yard6.push({
            value: this.rowexport_use[ax].primary_quantity6,
          });
        } else {
          this.all_result_yard6.push({
            value: 0,
          });
        }
      } else if (this.rowexport_use[ax].reason7 == "เสียจากแผนกแพด") {
        worksheet.getCell("AD " + [ax + 4]).value = Number(
          this.rowexport_use[ax].pcs7
        );
        worksheet.getCell("AD " + [ax + 4]).numFmt = "0";
        worksheet.getCell("AE " + [ax + 4]).value = Number(
          this.rowexport_use[ax].primary_quantity7
        );
        worksheet.getCell("AE " + [ax + 4]).numFmt = "0";
        if (this.rowexport_use[ax].pcs7 > 0) {
          this.all_result_pcs6.push({
            value: this.rowexport_use[ax].pcs7,
          });
        } else {
          this.all_result_pcs6.push({
            value: 0,
          });
        }

        if (this.rowexport_use[ax].primary_quantity7 > 0) {
          this.all_result_yard6.push({
            value: this.rowexport_use[ax].primary_quantity7,
          });
        } else {
          this.all_result_yard6.push({
            value: 0,
          });
        }
      } else if (this.rowexport_use[ax].reason8 == "เสียจากแผนกแพด") {
        worksheet.getCell("AD " + [ax + 4]).value = Number(
          this.rowexport_use[ax].pcs8
        );
        worksheet.getCell("AD " + [ax + 4]).numFmt = "0";
        worksheet.getCell("AE " + [ax + 4]).value = Number(
          this.rowexport_use[ax].primary_quantity8
        );
        worksheet.getCell("AE " + [ax + 4]).numFmt = "0";
        if (this.rowexport_use[ax].pcs8 > 0) {
          this.all_result_pcs6.push({
            value: this.rowexport_use[ax].pcs8,
          });
        } else {
          this.all_result_pcs6.push({
            value: 0,
          });
        }

        if (this.rowexport_use[ax].primary_quantity8 > 0) {
          this.all_result_yard6.push({
            value: this.rowexport_use[ax].primary_quantity8,
          });
        } else {
          this.all_result_yard6.push({
            value: 0,
          });
        }
      } else if (this.rowexport_use[ax].reason9 == "เสียจากแผนกแพด") {
        worksheet.getCell("AD " + [ax + 4]).value = Number(
          this.rowexport_use[ax].pcs9
        );
        worksheet.getCell("AD " + [ax + 4]).numFmt = "0";
        worksheet.getCell("AE " + [ax + 4]).value = Number(
          this.rowexport_use[ax].primary_quantity9
        );
        worksheet.getCell("AE " + [ax + 4]).numFmt = "0";
        if (this.rowexport_use[ax].pcs9 > 0) {
          this.all_result_pcs6.push({
            value: this.rowexport_use[ax].pcs9,
          });
        } else {
          this.all_result_pcs6.push({
            value: 0,
          });
        }

        if (this.rowexport_use[ax].primary_quantity9 > 0) {
          this.all_result_yard6.push({
            value: this.rowexport_use[ax].primary_quantity9,
          });
        } else {
          this.all_result_yard6.push({
            value: 0,
          });
        }
      } else if (this.rowexport_use[ax].reason10 == "เสียจากแผนกแพด") {
        worksheet.getCell("AD " + [ax + 4]).value = Number(
          this.rowexport_use[ax].pcs10
        );
        worksheet.getCell("AD " + [ax + 4]).numFmt = "0";
        worksheet.getCell("AE " + [ax + 4]).value = Number(
          this.rowexport_use[ax].primary_quantity10
        );
        worksheet.getCell("AE " + [ax + 4]).numFmt = "0";
        if (this.rowexport_use[ax].pcs10 > 0) {
          this.all_result_pcs6.push({
            value: this.rowexport_use[ax].pcs10,
          });
        } else {
          this.all_result_pcs6.push({
            value: 0,
          });
        }

        if (this.rowexport_use[ax].primary_quantity10 > 0) {
          this.all_result_yard6.push({
            value: this.rowexport_use[ax].primary_quantity10,
          });
        } else {
          this.all_result_yard6.push({
            value: 0,
          });
        }
      } else {
        worksheet.getCell("AD " + [ax + 4]).value = "-";
        worksheet.getCell("AE " + [ax + 4]).value = "-";
        this.all_result_pcs6.push({
          value: 0,
        });
        this.all_result_yard6.push({
          value: 0,
        });
      }
      if (this.rowexport_use[ax].reason == "เสียจากแผนกพิมพ์") {
        worksheet.getCell("AF " + [ax + 4]).value = Number(
          this.rowexport_use[ax].pcs
        );
        worksheet.getCell("AF " + [ax + 4]).numFmt = "0";
        worksheet.getCell("AG " + [ax + 4]).value = Number(
          this.rowexport_use[ax].primary_quantity
        );
        worksheet.getCell("AG " + [ax + 4]).numFmt = "0";
        if (this.rowexport_use[ax].pcs > 0) {
          this.all_result_pcs7.push({
            value: this.rowexport_use[ax].pcs,
          });
        } else {
          this.all_result_pcs7.push({
            value: 0,
          });
        }

        if (this.rowexport_use[ax].primary_quantity > 0) {
          this.all_result_yard7.push({
            value: this.rowexport_use[ax].primary_quantity,
          });
        } else {
          this.all_result_yard7.push({
            value: 0,
          });
        }
      } else if (this.rowexport_use[ax].reason2 == "เสียจากแผนกพิมพ์") {
        worksheet.getCell("AF " + [ax + 4]).value = Number(
          this.rowexport_use[ax].pcs2
        );
        worksheet.getCell("AF " + [ax + 4]).numFmt = "0";
        worksheet.getCell("AG " + [ax + 4]).value = Number(
          this.rowexport_use[ax].primary_quantity2
        );
        worksheet.getCell("AG " + [ax + 4]).numFmt = "0";
        if (this.rowexport_use[ax].pcs2 > 0) {
          this.all_result_pcs7.push({
            value: this.rowexport_use[ax].pcs2,
          });
        } else {
          this.all_result_pcs7.push({
            value: 0,
          });
        }

        if (this.rowexport_use[ax].primary_quantity2 > 0) {
          this.all_result_yard7.push({
            value: this.rowexport_use[ax].primary_quantity2,
          });
        } else {
          this.all_result_yard7.push({
            value: 0,
          });
        }
      } else if (this.rowexport_use[ax].reason3 == "เสียจากแผนกพิมพ์") {
        worksheet.getCell("AF " + [ax + 4]).value = Number(
          this.rowexport_use[ax].pcs3
        );
        worksheet.getCell("AF " + [ax + 4]).numFmt = "0";
        worksheet.getCell("AG " + [ax + 4]).value = Number(
          this.rowexport_use[ax].primary_quantity3
        );
        worksheet.getCell("AG " + [ax + 4]).numFmt = "0";
        if (this.rowexport_use[ax].pcs3 > 0) {
          this.all_result_pcs7.push({
            value: this.rowexport_use[ax].pcs3,
          });
        } else {
          this.all_result_pcs7.push({
            value: 0,
          });
        }

        if (this.rowexport_use[ax].primary_quantity3 > 0) {
          this.all_result_yard7.push({
            value: this.rowexport_use[ax].primary_quantity3,
          });
        } else {
          this.all_result_yard7.push({
            value: 0,
          });
        }
      } else if (this.rowexport_use[ax].reason4 == "เสียจากแผนกพิมพ์") {
        worksheet.getCell("AF " + [ax + 4]).value = Number(
          this.rowexport_use[ax].pcs4
        );
        worksheet.getCell("AF " + [ax + 4]).numFmt = "0";
        worksheet.getCell("AG " + [ax + 4]).value = Number(
          this.rowexport_use[ax].primary_quantity4
        );
        worksheet.getCell("AG " + [ax + 4]).numFmt = "0";
        if (this.rowexport_use[ax].pcs4 > 0) {
          this.all_result_pcs7.push({
            value: this.rowexport_use[ax].pcs4,
          });
        } else {
          this.all_result_pcs7.push({
            value: 0,
          });
        }

        if (this.rowexport_use[ax].primary_quantity4 > 0) {
          this.all_result_yard7.push({
            value: this.rowexport_use[ax].primary_quantity4,
          });
        } else {
          this.all_result_yard7.push({
            value: 0,
          });
        }
      } else if (this.rowexport_use[ax].reason5 == "เสียจากแผนกพิมพ์") {
        worksheet.getCell("AF " + [ax + 4]).value = Number(
          this.rowexport_use[ax].pcs5
        );
        worksheet.getCell("AF " + [ax + 4]).numFmt = "0";
        worksheet.getCell("AG " + [ax + 4]).value = Number(
          this.rowexport_use[ax].primary_quantity5
        );
        worksheet.getCell("AG " + [ax + 4]).numFmt = "0";
        if (this.rowexport_use[ax].pcs5 > 0) {
          this.all_result_pcs7.push({
            value: this.rowexport_use[ax].pcs5,
          });
        } else {
          this.all_result_pcs7.push({
            value: 0,
          });
        }

        if (this.rowexport_use[ax].primary_quantity5 > 0) {
          this.all_result_yard7.push({
            value: this.rowexport_use[ax].primary_quantity5,
          });
        } else {
          this.all_result_yard7.push({
            value: 0,
          });
        }
      } else if (this.rowexport_use[ax].reason6 == "เสียจากแผนกพิมพ์") {
        worksheet.getCell("AF " + [ax + 4]).value = Number(
          this.rowexport_use[ax].pcs6
        );
        worksheet.getCell("AF " + [ax + 4]).numFmt = "0";
        worksheet.getCell("AG " + [ax + 4]).value = Number(
          this.rowexport_use[ax].primary_quantity6
        );
        worksheet.getCell("AG " + [ax + 4]).numFmt = "0";
        if (this.rowexport_use[ax].pcs6 > 0) {
          this.all_result_pcs7.push({
            value: this.rowexport_use[ax].pcs6,
          });
        } else {
          this.all_result_pcs7.push({
            value: 0,
          });
        }

        if (this.rowexport_use[ax].primary_quantity6 > 0) {
          this.all_result_yard7.push({
            value: this.rowexport_use[ax].primary_quantity6,
          });
        } else {
          this.all_result_yard7.push({
            value: 0,
          });
        }
      } else if (this.rowexport_use[ax].reason7 == "เสียจากแผนกพิมพ์") {
        worksheet.getCell("AF " + [ax + 4]).value = Number(
          this.rowexport_use[ax].pcs7
        );
        worksheet.getCell("AF " + [ax + 4]).numFmt = "0";
        worksheet.getCell("AG " + [ax + 4]).value = Number(
          this.rowexport_use[ax].primary_quantity7
        );
        worksheet.getCell("AG " + [ax + 4]).numFmt = "0";
        if (this.rowexport_use[ax].pcs7 > 0) {
          this.all_result_pcs7.push({
            value: this.rowexport_use[ax].pcs7,
          });
        } else {
          this.all_result_pcs7.push({
            value: 0,
          });
        }

        if (this.rowexport_use[ax].primary_quantity7 > 0) {
          this.all_result_yard7.push({
            value: this.rowexport_use[ax].primary_quantity7,
          });
        } else {
          this.all_result_yard7.push({
            value: 0,
          });
        }
      } else if (this.rowexport_use[ax].reason8 == "เสียจากแผนกพิมพ์") {
        worksheet.getCell("AF " + [ax + 4]).value = Number(
          this.rowexport_use[ax].pcs8
        );
        worksheet.getCell("AF " + [ax + 4]).numFmt = "0";
        worksheet.getCell("AG " + [ax + 4]).value = Number(
          this.rowexport_use[ax].primary_quantity8
        );
        worksheet.getCell("AG " + [ax + 4]).numFmt = "0";
        if (this.rowexport_use[ax].pcs8 > 0) {
          this.all_result_pcs7.push({
            value: this.rowexport_use[ax].pcs8,
          });
        } else {
          this.all_result_pcs7.push({
            value: 0,
          });
        }

        if (this.rowexport_use[ax].primary_quantity8 > 0) {
          this.all_result_yard7.push({
            value: this.rowexport_use[ax].primary_quantity8,
          });
        } else {
          this.all_result_yard7.push({
            value: 0,
          });
        }
      } else if (this.rowexport_use[ax].reason9 == "เสียจากแผนกพิมพ์") {
        worksheet.getCell("AF " + [ax + 4]).value = Number(
          this.rowexport_use[ax].pcs9
        );
        worksheet.getCell("AF " + [ax + 4]).numFmt = "0";
        worksheet.getCell("AG " + [ax + 4]).value = Number(
          this.rowexport_use[ax].primary_quantity9
        );
        worksheet.getCell("AG " + [ax + 4]).numFmt = "0";
        if (this.rowexport_use[ax].pcs9 > 0) {
          this.all_result_pcs7.push({
            value: this.rowexport_use[ax].pcs9,
          });
        } else {
          this.all_result_pcs7.push({
            value: 0,
          });
        }

        if (this.rowexport_use[ax].primary_quantity9 > 0) {
          this.all_result_yard7.push({
            value: this.rowexport_use[ax].primary_quantity9,
          });
        } else {
          this.all_result_yard7.push({
            value: 0,
          });
        }
      } else if (this.rowexport_use[ax].reason10 == "เสียจากแผนกพิมพ์") {
        worksheet.getCell("AF " + [ax + 4]).value = Number(
          this.rowexport_use[ax].pcs10
        );
        worksheet.getCell("AF " + [ax + 4]).numFmt = "0";
        worksheet.getCell("AG " + [ax + 4]).value = Number(
          this.rowexport_use[ax].primary_quantity10
        );
        worksheet.getCell("AG " + [ax + 4]).numFmt = "0";
        if (this.rowexport_use[ax].pcs10 > 0) {
          this.all_result_pcs7.push({
            value: this.rowexport_use[ax].pcs10,
          });
        } else {
          this.all_result_pcs7.push({
            value: 0,
          });
        }

        if (this.rowexport_use[ax].primary_quantity10 > 0) {
          this.all_result_yard7.push({
            value: this.rowexport_use[ax].primary_quantity10,
          });
        } else {
          this.all_result_yard7.push({
            value: 0,
          });
        }
      } else {
        worksheet.getCell("AF " + [ax + 4]).value = "-";
        worksheet.getCell("AG " + [ax + 4]).value = "-";
        this.all_result_pcs7.push({
          value: 0,
        });
        this.all_result_yard7.push({
          value: 0,
        });
      }

      if (this.rowexport_use[ax].reason == "เสียจากแผนกฟิวส์") {
        worksheet.getCell("AH " + [ax + 4]).value = Number(
          this.rowexport_use[ax].pcs
        );
        worksheet.getCell("AH " + [ax + 4]).numFmt = "0";
        worksheet.getCell("AI " + [ax + 4]).value = Number(
          this.rowexport_use[ax].primary_quantity
        );
        worksheet.getCell("AI " + [ax + 4]).numFmt = "0";
        if (this.rowexport_use[ax].pcs > 0) {
          this.all_result_pcs8.push({
            value: this.rowexport_use[ax].pcs,
          });
        } else {
          this.all_result_pcs8.push({
            value: 0,
          });
        }

        if (this.rowexport_use[ax].primary_quantity > 0) {
          this.all_result_yard8.push({
            value: this.rowexport_use[ax].primary_quantity,
          });
        } else {
          this.all_result_yard8.push({
            value: 0,
          });
        }
      } else if (this.rowexport_use[ax].reason2 == "เสียจากแผนกฟิวส์") {
        worksheet.getCell("AH " + [ax + 4]).value = Number(
          this.rowexport_use[ax].pcs2
        );
        worksheet.getCell("AH " + [ax + 4]).numFmt = "0";
        worksheet.getCell("AI " + [ax + 4]).value = Number(
          this.rowexport_use[ax].primary_quantity2
        );
        worksheet.getCell("A " + [ax + 4]).numFmt = "0";
        if (this.rowexport_use[ax].pcs2 > 0) {
          this.all_result_pcs8.push({
            value: this.rowexport_use[ax].pcs2,
          });
        } else {
          this.all_result_pcs8.push({
            value: 0,
          });
        }

        if (this.rowexport_use[ax].primary_quantity2 > 0) {
          this.all_result_yard8.push({
            value: this.rowexport_use[ax].primary_quantity2,
          });
        } else {
          this.all_result_yard8.push({
            value: 0,
          });
        }
      } else if (this.rowexport_use[ax].reason3 == "เสียจากแผนกฟิวส์") {
        worksheet.getCell("AH " + [ax + 4]).value = Number(
          this.rowexport_use[ax].pcs3
        );
        worksheet.getCell("AF " + [ax + 4]).numFmt = "0";
        worksheet.getCell("AI " + [ax + 4]).value = Number(
          this.rowexport_use[ax].primary_quantity3
        );
        worksheet.getCell("AI " + [ax + 4]).numFmt = "0";
        if (this.rowexport_use[ax].pcs3 > 0) {
          this.all_result_pcs8.push({
            value: this.rowexport_use[ax].pcs3,
          });
        } else {
          this.all_result_pcs8.push({
            value: 0,
          });
        }

        if (this.rowexport_use[ax].primary_quantity3 > 0) {
          this.all_result_yard8.push({
            value: this.rowexport_use[ax].primary_quantity3,
          });
        } else {
          this.all_result_yard8.push({
            value: 0,
          });
        }
      } else if (this.rowexport_use[ax].reason4 == "เสียจากแผนกฟิวส์") {
        worksheet.getCell("AH " + [ax + 4]).value = Number(
          this.rowexport_use[ax].pcs4
        );
        worksheet.getCell("AH " + [ax + 4]).numFmt = "0";
        worksheet.getCell("AI " + [ax + 4]).value = Number(
          this.rowexport_use[ax].primary_quantity4
        );
        worksheet.getCell("AI " + [ax + 4]).numFmt = "0";
        if (this.rowexport_use[ax].pcs4 > 0) {
          this.all_result_pcs8.push({
            value: this.rowexport_use[ax].pcs4,
          });
        } else {
          this.all_result_pcs8.push({
            value: 0,
          });
        }

        if (this.rowexport_use[ax].primary_quantity4 > 0) {
          this.all_result_yard8.push({
            value: this.rowexport_use[ax].primary_quantity4,
          });
        } else {
          this.all_result_yard8.push({
            value: 0,
          });
        }
      } else if (this.rowexport_use[ax].reason5 == "เสียจากแผนกฟิวส์") {
        worksheet.getCell("AH " + [ax + 4]).value = Number(
          this.rowexport_use[ax].pcs5
        );
        worksheet.getCell("AH " + [ax + 4]).numFmt = "0";
        worksheet.getCell("AI " + [ax + 4]).value = Number(
          this.rowexport_use[ax].primary_quantity5
        );
        worksheet.getCell("AI " + [ax + 4]).numFmt = "0";
        if (this.rowexport_use[ax].pcs5 > 0) {
          this.all_result_pcs8.push({
            value: this.rowexport_use[ax].pcs5,
          });
        } else {
          this.all_result_pcs8.push({
            value: 0,
          });
        }

        if (this.rowexport_use[ax].primary_quantity5 > 0) {
          this.all_result_yard8.push({
            value: this.rowexport_use[ax].primary_quantity5,
          });
        } else {
          this.all_result_yard8.push({
            value: 0,
          });
        }
      } else if (this.rowexport_use[ax].reason6 == "เสียจากแผนกฟิวส์") {
        worksheet.getCell("AH " + [ax + 4]).value = Number(
          this.rowexport_use[ax].pcs6
        );
        worksheet.getCell("AH " + [ax + 4]).numFmt = "0";
        worksheet.getCell("AI " + [ax + 4]).value = Number(
          this.rowexport_use[ax].primary_quantity6
        );
        worksheet.getCell("AI " + [ax + 4]).numFmt = "0";
        if (this.rowexport_use[ax].pcs6 > 0) {
          this.all_result_pcs8.push({
            value: this.rowexport_use[ax].pcs6,
          });
        } else {
          this.all_result_pcs8.push({
            value: 0,
          });
        }

        if (this.rowexport_use[ax].primary_quantity6 > 0) {
          this.all_result_yard8.push({
            value: this.rowexport_use[ax].primary_quantity6,
          });
        } else {
          this.all_result_yard8.push({
            value: 0,
          });
        }
      } else if (this.rowexport_use[ax].reason7 == "เสียจากแผนกฟิวส์") {
        worksheet.getCell("AH " + [ax + 4]).value = Number(
          this.rowexport_use[ax].pcs7
        );
        worksheet.getCell("AH " + [ax + 4]).numFmt = "0";
        worksheet.getCell("AI " + [ax + 4]).value = Number(
          this.rowexport_use[ax].primary_quantity7
        );
        worksheet.getCell("AI " + [ax + 4]).numFmt = "0";
        if (this.rowexport_use[ax].pcs7 > 0) {
          this.all_result_pcs8.push({
            value: this.rowexport_use[ax].pcs7,
          });
        } else {
          this.all_result_pcs8.push({
            value: 0,
          });
        }

        if (this.rowexport_use[ax].primary_quantity7 > 0) {
          this.all_result_yard8.push({
            value: this.rowexport_use[ax].primary_quantity7,
          });
        } else {
          this.all_result_yard8.push({
            value: 0,
          });
        }
      } else if (this.rowexport_use[ax].reason8 == "เสียจากแผนกฟิวส์") {
        worksheet.getCell("AH " + [ax + 4]).value = Number(
          this.rowexport_use[ax].pcs8
        );
        worksheet.getCell("AH " + [ax + 4]).numFmt = "0";
        worksheet.getCell("AI " + [ax + 4]).value = Number(
          this.rowexport_use[ax].primary_quantity8
        );
        worksheet.getCell("AI " + [ax + 4]).numFmt = "0";
        if (this.rowexport_use[ax].pcs8 > 0) {
          this.all_result_pcs8.push({
            value: this.rowexport_use[ax].pcs8,
          });
        } else {
          this.all_result_pcs8.push({
            value: 0,
          });
        }

        if (this.rowexport_use[ax].primary_quantity8 > 0) {
          this.all_result_yard8.push({
            value: this.rowexport_use[ax].primary_quantity8,
          });
        } else {
          this.all_result_yard8.push({
            value: 0,
          });
        }
      } else if (this.rowexport_use[ax].reason9 == "เสียจากแผนกฟิวส์") {
        worksheet.getCell("AH " + [ax + 4]).value = Number(
          this.rowexport_use[ax].pcs9
        );
        worksheet.getCell("AH " + [ax + 4]).numFmt = "0";
        worksheet.getCell("AI " + [ax + 4]).value = Number(
          this.rowexport_use[ax].primary_quantity9
        );
        worksheet.getCell("AI " + [ax + 4]).numFmt = "0";
        if (this.rowexport_use[ax].pcs9 > 0) {
          this.all_result_pcs8.push({
            value: this.rowexport_use[ax].pcs9,
          });
        } else {
          this.all_result_pcs8.push({
            value: 0,
          });
        }

        if (this.rowexport_use[ax].primary_quantity9 > 0) {
          this.all_result_yard8.push({
            value: this.rowexport_use[ax].primary_quantity9,
          });
        } else {
          this.all_result_yard8.push({
            value: 0,
          });
        }
      } else if (this.rowexport_use[ax].reason10 == "เสียจากแผนกฟิวส์") {
        worksheet.getCell("AH " + [ax + 4]).value = Number(
          this.rowexport_use[ax].pcs10
        );
        worksheet.getCell("AH " + [ax + 4]).numFmt = "0";
        worksheet.getCell("AI " + [ax + 4]).value = Number(
          this.rowexport_use[ax].primary_quantity10
        );
        worksheet.getCell("AI " + [ax + 4]).numFmt = "0";
        if (this.rowexport_use[ax].pcs10 > 0) {
          this.all_result_pcs8.push({
            value: this.rowexport_use[ax].pcs10,
          });
        } else {
          this.all_result_pcs8.push({
            value: 0,
          });
        }

        if (this.rowexport_use[ax].primary_quantity10 > 0) {
          this.all_result_yard8.push({
            value: this.rowexport_use[ax].primary_quantity10,
          });
        } else {
          this.all_result_yard8.push({
            value: 0,
          });
        }
      } else {
        worksheet.getCell("AH " + [ax + 4]).value = "-";
        worksheet.getCell("AI " + [ax + 4]).value = "-";
        this.all_result_pcs8.push({
          value: 0,
        });
        this.all_result_yard8.push({
          value: 0,
        });
      }

      if (this.rowexport_use[ax].reason == "เสียจากแผนกปัก") {
        worksheet.getCell("AJ " + [ax + 4]).value = Number(
          this.rowexport_use[ax].pcs
        );
        worksheet.getCell("AJ " + [ax + 4]).numFmt = "0";
        worksheet.getCell("AK " + [ax + 4]).value = Number(
          this.rowexport_use[ax].primary_quantity
        );
        worksheet.getCell("AK " + [ax + 4]).numFmt = "0";
        if (this.rowexport_use[ax].pcs > 0) {
          this.all_result_pcs9.push({
            value: this.rowexport_use[ax].pcs,
          });
        } else {
          this.all_result_pcs9.push({
            value: 0,
          });
        }

        if (this.rowexport_use[ax].primary_quantity > 0) {
          this.all_result_yard9.push({
            value: this.rowexport_use[ax].primary_quantity,
          });
        } else {
          this.all_result_yard9.push({
            value: 0,
          });
        }
      } else if (this.rowexport_use[ax].reason2 == "เสียจากแผนกปัก") {
        worksheet.getCell("AJ " + [ax + 4]).value = Number(
          this.rowexport_use[ax].pcs2
        );
        worksheet.getCell("AJ " + [ax + 4]).numFmt = "0";
        worksheet.getCell("AK " + [ax + 4]).value = Number(
          this.rowexport_use[ax].primary_quantity2
        );
        worksheet.getCell("AK " + [ax + 4]).numFmt = "0";
        if (this.rowexport_use[ax].pcs2 > 0) {
          this.all_result_pcs9.push({
            value: this.rowexport_use[ax].pcs2,
          });
        } else {
          this.all_result_pcs9.push({
            value: 0,
          });
        }

        if (this.rowexport_use[ax].primary_quantity2 > 0) {
          this.all_result_yard9.push({
            value: this.rowexport_use[ax].primary_quantity2,
          });
        } else {
          this.all_result_yard9.push({
            value: 0,
          });
        }
      } else if (this.rowexport_use[ax].reason3 == "เสียจากแผนกปัก") {
        worksheet.getCell("AJ " + [ax + 4]).value = Number(
          this.rowexport_use[ax].pcs3
        );
        worksheet.getCell("AJ " + [ax + 4]).numFmt = "0";
        worksheet.getCell("AK " + [ax + 4]).value = Number(
          this.rowexport_use[ax].primary_quantity3
        );
        worksheet.getCell("AK " + [ax + 4]).numFmt = "0";
        if (this.rowexport_use[ax].pcs3 > 0) {
          this.all_result_pcs9.push({
            value: this.rowexport_use[ax].pcs3,
          });
        } else {
          this.all_result_pcs9.push({
            value: 0,
          });
        }

        if (this.rowexport_use[ax].primary_quantity3 > 0) {
          this.all_result_yard9.push({
            value: this.rowexport_use[ax].primary_quantity3,
          });
        } else {
          this.all_result_yard9.push({
            value: 0,
          });
        }
      } else if (this.rowexport_use[ax].reason4 == "เสียจากแผนกปัก") {
        worksheet.getCell("AJ " + [ax + 4]).value = Number(
          this.rowexport_use[ax].pcs4
        );
        worksheet.getCell("AJ " + [ax + 4]).numFmt = "0";
        worksheet.getCell("AK " + [ax + 4]).value = Number(
          this.rowexport_use[ax].primary_quantity4
        );
        worksheet.getCell("AK " + [ax + 4]).numFmt = "0";
        if (this.rowexport_use[ax].pcs4 > 0) {
          this.all_result_pcs9.push({
            value: this.rowexport_use[ax].pcs4,
          });
        } else {
          this.all_result_pcs9.push({
            value: 0,
          });
        }

        if (this.rowexport_use[ax].primary_quantity4 > 0) {
          this.all_result_yard9.push({
            value: this.rowexport_use[ax].primary_quantity4,
          });
        } else {
          this.all_result_yard9.push({
            value: 0,
          });
        }
      } else if (this.rowexport_use[ax].reason5 == "เสียจากแผนกปัก") {
        worksheet.getCell("AJ " + [ax + 4]).value = Number(
          this.rowexport_use[ax].pcs5
        );
        worksheet.getCell("AJ " + [ax + 4]).numFmt = "0";
        worksheet.getCell("AK " + [ax + 4]).value = Number(
          this.rowexport_use[ax].primary_quantity5
        );
        worksheet.getCell("AK " + [ax + 4]).numFmt = "0";
        if (this.rowexport_use[ax].pcs5 > 0) {
          this.all_result_pcs9.push({
            value: this.rowexport_use[ax].pcs5,
          });
        } else {
          this.all_result_pcs9.push({
            value: 0,
          });
        }

        if (this.rowexport_use[ax].primary_quantity5 > 0) {
          this.all_result_yard9.push({
            value: this.rowexport_use[ax].primary_quantity5,
          });
        } else {
          this.all_result_yard9.push({
            value: 0,
          });
        }
      } else if (this.rowexport_use[ax].reason6 == "เสียจากแผนกปัก") {
        worksheet.getCell("AJ " + [ax + 4]).value = Number(
          this.rowexport_use[ax].pcs6
        );
        worksheet.getCell("AJ " + [ax + 4]).numFmt = "0";
        worksheet.getCell("AK " + [ax + 4]).value = Number(
          this.rowexport_use[ax].primary_quantity6
        );
        worksheet.getCell("AK " + [ax + 4]).numFmt = "0";
        if (this.rowexport_use[ax].pcs6 > 0) {
          this.all_result_pcs9.push({
            value: this.rowexport_use[ax].pcs6,
          });
        } else {
          this.all_result_pcs9.push({
            value: 0,
          });
        }

        if (this.rowexport_use[ax].primary_quantity6 > 0) {
          this.all_result_yard9.push({
            value: this.rowexport_use[ax].primary_quantity6,
          });
        } else {
          this.all_result_yard9.push({
            value: 0,
          });
        }
      } else if (this.rowexport_use[ax].reason7 == "เสียจากแผนกปัก") {
        worksheet.getCell("AJ " + [ax + 4]).value = Number(
          this.rowexport_use[ax].pcs7
        );
        worksheet.getCell("AJ " + [ax + 4]).numFmt = "0";
        worksheet.getCell("AK " + [ax + 4]).value = Number(
          this.rowexport_use[ax].primary_quantity7
        );
        worksheet.getCell("AK " + [ax + 4]).numFmt = "0";
        if (this.rowexport_use[ax].pcs7 > 0) {
          this.all_result_pcs9.push({
            value: this.rowexport_use[ax].pcs7,
          });
        } else {
          this.all_result_pcs9.push({
            value: 0,
          });
        }

        if (this.rowexport_use[ax].primary_quantity7 > 0) {
          this.all_result_yard9.push({
            value: this.rowexport_use[ax].primary_quantity7,
          });
        } else {
          this.all_result_yard9.push({
            value: 0,
          });
        }
      } else if (this.rowexport_use[ax].reason8 == "เสียจากแผนกปัก") {
        worksheet.getCell("AJ " + [ax + 4]).value = Number(
          this.rowexport_use[ax].pcs8
        );
        worksheet.getCell("AJ " + [ax + 4]).numFmt = "0";
        worksheet.getCell("AK " + [ax + 4]).value = Number(
          this.rowexport_use[ax].primary_quantity8
        );
        worksheet.getCell("AK " + [ax + 4]).numFmt = "0";
        if (this.rowexport_use[ax].pcs8 > 0) {
          this.all_result_pcs9.push({
            value: this.rowexport_use[ax].pcs8,
          });
        } else {
          this.all_result_pcs9.push({
            value: 0,
          });
        }

        if (this.rowexport_use[ax].primary_quantity8 > 0) {
          this.all_result_yard9.push({
            value: this.rowexport_use[ax].primary_quantity8,
          });
        } else {
          this.all_result_yard9.push({
            value: 0,
          });
        }
      } else if (this.rowexport_use[ax].reason9 == "เสียจากแผนกปัก") {
        worksheet.getCell("AJ " + [ax + 4]).value = Number(
          this.rowexport_use[ax].pcs9
        );
        worksheet.getCell("AJ " + [ax + 4]).numFmt = "0";
        worksheet.getCell("AK " + [ax + 4]).value = Number(
          this.rowexport_use[ax].primary_quantity9
        );
        worksheet.getCell("AK " + [ax + 4]).numFmt = "0";
        if (this.rowexport_use[ax].pcs9 > 0) {
          this.all_result_pcs9.push({
            value: this.rowexport_use[ax].pcs9,
          });
        } else {
          this.all_result_pcs9.push({
            value: 0,
          });
        }

        if (this.rowexport_use[ax].primary_quantity9 > 0) {
          this.all_result_yard9.push({
            value: this.rowexport_use[ax].primary_quantity9,
          });
        } else {
          this.all_result_yard9.push({
            value: 0,
          });
        }
      } else if (this.rowexport_use[ax].reason10 == "เสียจากแผนกปัก") {
        worksheet.getCell("AJ " + [ax + 4]).value = Number(
          this.rowexport_use[ax].pcs10
        );
        worksheet.getCell("AJ " + [ax + 4]).numFmt = "0";
        worksheet.getCell("AK " + [ax + 4]).value = Number(
          this.rowexport_use[ax].primary_quantity10
        );
        worksheet.getCell("AK " + [ax + 4]).numFmt = "0";
        if (this.rowexport_use[ax].pcs10 > 0) {
          this.all_result_pcs9.push({
            value: this.rowexport_use[ax].pcs10,
          });
        } else {
          this.all_result_pcs9.push({
            value: 0,
          });
        }

        if (this.rowexport_use[ax].primary_quantity10 > 0) {
          this.all_result_yard9.push({
            value: this.rowexport_use[ax].primary_quantity10,
          });
        } else {
          this.all_result_yard9.push({
            value: 0,
          });
        }
      } else {
        worksheet.getCell("AJ " + [ax + 4]).value = "-";
        worksheet.getCell("AK " + [ax + 4]).value = "-";
        this.all_result_pcs9.push({
          value: 0,
        });
        this.all_result_yard9.push({
          value: 0,
        });
      }

      if (this.rowexport_use[ax].reason == "งานหาย") {
        worksheet.getCell("AL " + [ax + 4]).value = Number(
          this.rowexport_use[ax].pcs
        );
        worksheet.getCell("AL " + [ax + 4]).numFmt = "0";
        worksheet.getCell("AM " + [ax + 4]).value = Number(
          this.rowexport_use[ax].primary_quantity
        );
        worksheet.getCell("AM " + [ax + 4]).numFmt = "0";
        if (this.rowexport_use[ax].pcs > 0) {
          this.all_result_pcs10.push({
            value: this.rowexport_use[ax].pcs,
          });
        } else {
          this.all_result_pcs10.push({
            value: 0,
          });
        }

        if (this.rowexport_use[ax].primary_quantity > 0) {
          this.all_result_yard10.push({
            value: this.rowexport_use[ax].primary_quantity,
          });
        } else {
          this.all_result_yard10.push({
            value: 0,
          });
        }
      } else if (this.rowexport_use[ax].reason2 == "งานหาย") {
        worksheet.getCell("AL " + [ax + 4]).value = Number(
          this.rowexport_use[ax].pcs2
        );
        worksheet.getCell("AL " + [ax + 4]).numFmt = "0";
        worksheet.getCell("AM " + [ax + 4]).value = Number(
          this.rowexport_use[ax].primary_quantity2
        );
        worksheet.getCell("AM " + [ax + 4]).numFmt = "0";
        if (this.rowexport_use[ax].pcs2 > 0) {
          this.all_result_pcs10.push({
            value: this.rowexport_use[ax].pcs2,
          });
        } else {
          this.all_result_pcs10.push({
            value: 0,
          });
        }

        if (this.rowexport_use[ax].primary_quantity2 > 0) {
          this.all_result_yard10.push({
            value: this.rowexport_use[ax].primary_quantity2,
          });
        } else {
          this.all_result_yard10.push({
            value: 0,
          });
        }
      } else if (this.rowexport_use[ax].reason3 == "งานหาย") {
        worksheet.getCell("AL " + [ax + 4]).value = Number(
          this.rowexport_use[ax].pcs3
        );
        worksheet.getCell("AL " + [ax + 4]).numFmt = "0";
        worksheet.getCell("AM " + [ax + 4]).value = Number(
          this.rowexport_use[ax].primary_quantity3
        );
        worksheet.getCell("AM " + [ax + 4]).numFmt = "0";
        if (this.rowexport_use[ax].pcs3 > 0) {
          this.all_result_pcs10.push({
            value: this.rowexport_use[ax].pcs3,
          });
        } else {
          this.all_result_pcs.push({
            value: 0,
          });
        }

        if (this.rowexport_use[ax].primary_quantity3 > 0) {
          this.all_result_yard10.push({
            value: this.rowexport_use[ax].primary_quantity3,
          });
        } else {
          this.all_result_yard10.push({
            value: 0,
          });
        }
      } else if (this.rowexport_use[ax].reason4 == "งานหาย") {
        worksheet.getCell("AL " + [ax + 4]).value = Number(
          this.rowexport_use[ax].pcs4
        );
        worksheet.getCell("AL " + [ax + 4]).numFmt = "0";
        worksheet.getCell("AM " + [ax + 4]).value = Number(
          this.rowexport_use[ax].primary_quantity4
        );
        worksheet.getCell("AM " + [ax + 4]).numFmt = "0";
        if (this.rowexport_use[ax].pcs4 > 0) {
          this.all_result_pcs10.push({
            value: this.rowexport_use[ax].pcs4,
          });
        } else {
          this.all_result_pcs10.push({
            value: 0,
          });
        }

        if (this.rowexport_use[ax].primary_quantity4 > 0) {
          this.all_result_yard10.push({
            value: this.rowexport_use[ax].primary_quantity4,
          });
        } else {
          this.all_result_yard10.push({
            value: 0,
          });
        }
      } else if (this.rowexport_use[ax].reason5 == "งานหาย") {
        worksheet.getCell("AL " + [ax + 4]).value = Number(
          this.rowexport_use[ax].pcs5
        );
        worksheet.getCell("AL " + [ax + 4]).numFmt = "0";
        worksheet.getCell("AM " + [ax + 4]).value = Number(
          this.rowexport_use[ax].primary_quantity5
        );
        worksheet.getCell("AM " + [ax + 4]).numFmt = "0";
        if (this.rowexport_use[ax].pcs5 > 0) {
          this.all_result_pcs10.push({
            value: this.rowexport_use[ax].pcs5,
          });
        } else {
          this.all_result_pcs10.push({
            value: 0,
          });
        }

        if (this.rowexport_use[ax].primary_quantity5 > 0) {
          this.all_result_yard10.push({
            value: this.rowexport_use[ax].primary_quantity5,
          });
        } else {
          this.all_result_yard10.push({
            value: 0,
          });
        }
      } else if (this.rowexport_use[ax].reason6 == "งานหาย") {
        worksheet.getCell("AL " + [ax + 4]).value = Number(
          this.rowexport_use[ax].pcs6
        );
        worksheet.getCell("AL " + [ax + 4]).numFmt = "0";
        worksheet.getCell("AM " + [ax + 4]).value = Number(
          this.rowexport_use[ax].primary_quantity6
        );
        worksheet.getCell("AM " + [ax + 4]).numFmt = "0";
        if (this.rowexport_use[ax].pcs6 > 0) {
          this.all_result_pcs10.push({
            value: this.rowexport_use[ax].pcs6,
          });
        } else {
          this.all_result_pcs10.push({
            value: 0,
          });
        }

        if (this.rowexport_use[ax].primary_quantity6 > 0) {
          this.all_result_yard10.push({
            value: this.rowexport_use[ax].primary_quantity6,
          });
        } else {
          this.all_result_yard10.push({
            value: 0,
          });
        }
      } else if (this.rowexport_use[ax].reason7 == "งานหาย") {
        worksheet.getCell("AL " + [ax + 4]).value = Number(
          this.rowexport_use[ax].pcs7
        );
        worksheet.getCell("AL " + [ax + 4]).numFmt = "0";
        worksheet.getCell("AM " + [ax + 4]).value = Number(
          this.rowexport_use[ax].primary_quantity7
        );
        worksheet.getCell("AM " + [ax + 4]).numFmt = "0";
        if (this.rowexport_use[ax].pcs7 > 0) {
          this.all_result_pcs10.push({
            value: this.rowexport_use[ax].pcs7,
          });
        } else {
          this.all_result_pcs10.push({
            value: 0,
          });
        }

        if (this.rowexport_use[ax].primary_quantity7 > 0) {
          this.all_result_yard10.push({
            value: this.rowexport_use[ax].primary_quantity7,
          });
        } else {
          this.all_result_yard10.push({
            value: 0,
          });
        }
      } else if (this.rowexport_use[ax].reason8 == "งานหาย") {
        worksheet.getCell("AL " + [ax + 4]).value = Number(
          this.rowexport_use[ax].pcs8
        );
        worksheet.getCell("AL " + [ax + 4]).numFmt = "0";
        worksheet.getCell("AM " + [ax + 4]).value = Number(
          this.rowexport_use[ax].primary_quantity8
        );
        worksheet.getCell("AM " + [ax + 4]).numFmt = "0";
        if (this.rowexport_use[ax].pcs8 > 0) {
          this.all_result_pcs10.push({
            value: this.rowexport_use[ax].pcs8,
          });
        } else {
          this.all_result_pcs10.push({
            value: 0,
          });
        }

        if (this.rowexport_use[ax].primary_quantity8 > 0) {
          this.all_result_yard10.push({
            value: this.rowexport_use[ax].primary_quantity8,
          });
        } else {
          this.all_result_yard10.push({
            value: 0,
          });
        }
      } else if (this.rowexport_use[ax].reason9 == "งานหาย") {
        worksheet.getCell("AL " + [ax + 4]).value = Number(
          this.rowexport_use[ax].pcs9
        );
        worksheet.getCell("AL " + [ax + 4]).numFmt = "0";
        worksheet.getCell("AM " + [ax + 4]).value = Number(
          this.rowexport_use[ax].primary_quantity9
        );
        worksheet.getCell("AM " + [ax + 4]).numFmt = "0";
        if (this.rowexport_use[ax].pcs9 > 0) {
          this.all_result_pcs10.push({
            value: this.rowexport_use[ax].pcs9,
          });
        } else {
          this.all_result_pcs10.push({
            value: 0,
          });
        }

        if (this.rowexport_use[ax].primary_quantity9 > 0) {
          this.all_result_yard10.push({
            value: this.rowexport_use[ax].primary_quantity9,
          });
        } else {
          this.all_result_yard10.push({
            value: 0,
          });
        }
      } else if (this.rowexport_use[ax].reason10 == "งานหาย") {
        worksheet.getCell("AL " + [ax + 4]).value = Number(
          this.rowexport_use[ax].pcs10
        );
        worksheet.getCell("AL " + [ax + 4]).numFmt = "0";
        worksheet.getCell("AM " + [ax + 4]).value = Number(
          this.rowexport_use[ax].primary_quantity10
        );
        worksheet.getCell("AM " + [ax + 4]).numFmt = "0";
        if (this.rowexport_use[ax].pcs10 > 0) {
          this.all_result_pcs10.push({
            value: this.rowexport_use[ax].pcs10,
          });
        } else {
          this.all_result_pcs10.push({
            value: 0,
          });
        }

        if (this.rowexport_use[ax].primary_quantity10 > 0) {
          this.all_result_yard10.push({
            value: this.rowexport_use[ax].primary_quantity10,
          });
        } else {
          this.all_result_yard10.push({
            value: 0,
          });
        }
      } else {
        worksheet.getCell("AL " + [ax + 4]).value = "-";
        worksheet.getCell("AM " + [ax + 4]).value = "-";
        this.all_result_pcs10.push({
          value: 0,
        });
        this.all_result_yard10.push({
          value: 0,
        });
      }
      if (this.rowexport_use[ax].reason == "NA") {
        worksheet.getCell("AR " + [ax + 4]).value = Number(
          this.rowexport_use[ax].pcs
        );
        worksheet.getCell("AR " + [ax + 4]).numFmt = "0";
        worksheet.getCell("AS " + [ax + 4]).value = Number(
          this.rowexport_use[ax].primary_quantity
        );
        worksheet.getCell("AS " + [ax + 4]).numFmt = "0";
        if (this.rowexport_use[ax].pcs > 0) {
          this.all_result_pcs11.push({
            value: this.rowexport_use[ax].pcs,
          });
        } else {
          this.all_result_pcs11.push({
            value: 0,
          });
        }

        if (this.rowexport_use[ax].primary_quantity > 0) {
          this.all_result_yard11.push({
            value: this.rowexport_use[ax].primary_quantity,
          });
        } else {
          this.all_result_yard11.push({
            value: 0,
          });
        }
      } else if (this.rowexport_use[ax].reason2 == "NA") {
        worksheet.getCell("AR " + [ax + 4]).value = Number(
          this.rowexport_use[ax].pcs2
        );
        worksheet.getCell("AR " + [ax + 4]).numFmt = "0";
        worksheet.getCell("AS " + [ax + 4]).value = Number(
          this.rowexport_use[ax].primary_quantity2
        );
        worksheet.getCell("AS " + [ax + 4]).numFmt = "0";
        if (this.rowexport_use[ax].pcs2 > 0) {
          this.all_result_pcs11.push({
            value: this.rowexport_use[ax].pcs2,
          });
        } else {
          this.all_result_pcs11.push({
            value: 0,
          });
        }

        if (this.rowexport_use[ax].primary_quantity2 > 0) {
          this.all_result_yard11.push({
            value: this.rowexport_use[ax].primary_quantity2,
          });
        } else {
          this.all_result_yard11.push({
            value: 0,
          });
        }
      } else if (this.rowexport_use[ax].reason3 == "NA") {
        worksheet.getCell("AR " + [ax + 4]).value = Number(
          this.rowexport_use[ax].pcs3
        );
        worksheet.getCell("AR " + [ax + 4]).numFmt = "0";
        worksheet.getCell("AS " + [ax + 4]).value = Number(
          this.rowexport_use[ax].primary_quantity3
        );
        worksheet.getCell("AS " + [ax + 4]).numFmt = "0";
        if (this.rowexport_use[ax].result_sum_total_pcs3 > 0) {
          this.all_result_pcs11.push({
            value: this.rowexport_use[ax].pcs3,
          });
        } else {
          this.all_result_pcs11.push({
            value: 0,
          });
        }

        if (this.rowexport_use[ax].primary_quantity3 > 0) {
          this.all_result_yard11.push({
            value: this.rowexport_use[ax].primary_quantity3,
          });
        } else {
          this.all_result_yard11.push({
            value: 0,
          });
        }
      } else if (this.rowexport_use[ax].reason4 == "NA") {
        worksheet.getCell("AR " + [ax + 4]).value = Number(
          this.rowexport_use[ax].pcs4
        );
        worksheet.getCell("AR " + [ax + 4]).numFmt = "0";
        worksheet.getCell("AS " + [ax + 4]).value = Number(
          this.rowexport_use[ax].primary_quantity4
        );
        worksheet.getCell("AS " + [ax + 4]).numFmt = "0";
        if (this.rowexport_use[ax].pcs4 > 0) {
          this.all_result_pcs11.push({
            value: this.rowexport_use[ax].pcs4,
          });
        } else {
          this.all_result_pcs11.push({
            value: 0,
          });
        }

        if (this.rowexport_use[ax].primary_quantity4 > 0) {
          this.all_result_yard11.push({
            value: this.rowexport_use[ax].primary_quantity4,
          });
        } else {
          this.all_result_yard11.push({
            value: 0,
          });
        }
      } else if (this.rowexport_use[ax].reason5 == "NA") {
        worksheet.getCell("AR " + [ax + 4]).value = Number(
          this.rowexport_use[ax].pcs5
        );
        worksheet.getCell("AR " + [ax + 4]).numFmt = "0";
        worksheet.getCell("AS " + [ax + 4]).value = Number(
          this.rowexport_use[ax].primary_quantity5
        );
        worksheet.getCell("AS " + [ax + 4]).numFmt = "0";
        if (this.rowexport_use[ax].pcs5 > 0) {
          this.all_result_pcs11.push({
            value: this.rowexport_use[ax].pcs5,
          });
        } else {
          this.all_result_pcs11.push({
            value: 0,
          });
        }

        if (this.rowexport_use[ax].primary_quantity5 > 0) {
          this.all_result_yard11.push({
            value: this.rowexport_use[ax].primary_quantity5,
          });
        } else {
          this.all_result_yard11.push({
            value: 0,
          });
        }
      } else if (this.rowexport_use[ax].reason6 == "NA") {
        worksheet.getCell("AR " + [ax + 4]).value = Number(
          this.rowexport_use[ax].pcs6
        );
        worksheet.getCell("AR " + [ax + 4]).numFmt = "0";
        worksheet.getCell("AS " + [ax + 4]).value = Number(
          this.rowexport_use[ax].primary_quantity6
        );
        worksheet.getCell("AS " + [ax + 4]).numFmt = "0";
        if (this.rowexport_use[ax].pcs6 > 0) {
          this.all_result_pcs11.push({
            value: this.rowexport_use[ax].pcs6,
          });
        } else {
          this.all_result_pcs11.push({
            value: 0,
          });
        }

        if (this.rowexport_use[ax].primary_quantity6 > 0) {
          this.all_result_yard11.push({
            value: this.rowexport_use[ax].primary_quantity6,
          });
        } else {
          this.all_result_yard11.push({
            value: 0,
          });
        }
      } else if (this.rowexport_use[ax].reason7 == "NA") {
        worksheet.getCell("AR " + [ax + 4]).value = Number(
          this.rowexport_use[ax].pcs7
        );
        worksheet.getCell("AR " + [ax + 4]).numFmt = "0";
        worksheet.getCell("AS " + [ax + 4]).value = Number(
          this.rowexport_use[ax].primary_quantity7
        );
        worksheet.getCell("AS " + [ax + 4]).numFmt = "0";
        if (this.rowexport_use[ax].pcs7 > 0) {
          this.all_result_pcs11.push({
            value: this.rowexport_use[ax].pcs7,
          });
        } else {
          this.all_result_pcs11.push({
            value: 0,
          });
        }

        if (this.rowexport_use[ax].primary_quantity7 > 0) {
          this.all_result_yard11.push({
            value: this.rowexport_use[ax].primary_quantity7,
          });
        } else {
          this.all_result_yard11.push({
            value: 0,
          });
        }
      } else if (this.rowexport_use[ax].reason8 == "NA") {
        worksheet.getCell("AR " + [ax + 4]).value = Number(
          this.rowexport_use[ax].pcs8
        );
        worksheet.getCell("AR " + [ax + 4]).numFmt = "0";
        worksheet.getCell("AS " + [ax + 4]).value = Number(
          this.rowexport_use[ax].primary_quantity8
        );
        worksheet.getCell("AS " + [ax + 4]).numFmt = "0";
        if (this.rowexport_use[ax].pcs8 > 0) {
          this.all_result_pcs11.push({
            value: this.rowexport_use[ax].pcs8,
          });
        } else {
          this.all_result_pcs11.push({
            value: 0,
          });
        }

        if (this.rowexport_use[ax].primary_quantity8 > 0) {
          this.all_result_yard11.push({
            value: this.rowexport_use[ax].primary_quantity8,
          });
        } else {
          this.all_result_yard11.push({
            value: 0,
          });
        }
      } else if (this.rowexport_use[ax].reason9 == "NA") {
        worksheet.getCell("AR " + [ax + 4]).value = Number(
          this.rowexport_use[ax].pcs9
        );
        worksheet.getCell("AR " + [ax + 4]).numFmt = "0";
        worksheet.getCell("AS " + [ax + 4]).value = Number(
          this.rowexport_use[ax].primary_quantity9
        );
        worksheet.getCell("AS " + [ax + 4]).numFmt = "0";
        if (this.rowexport_use[ax].pcs9 > 0) {
          this.all_result_pcs11.push({
            value: this.rowexport_use[ax].pcs9,
          });
        } else {
          this.all_result_pcs11.push({
            value: 0,
          });
        }

        if (this.rowexport_use[ax].primary_quantity9 > 0) {
          this.all_result_yard11.push({
            value: this.rowexport_use[ax].primary_quantity9,
          });
        } else {
          this.all_result_yard11.push({
            value: 0,
          });
        }
      } else if (this.rowexport_use[ax].reason10 == "NA") {
        worksheet.getCell("AR " + [ax + 4]).value = Number(
          this.rowexport_use[ax].pcs10
        );
        worksheet.getCell("AR " + [ax + 4]).numFmt = "0";
        worksheet.getCell("AS " + [ax + 4]).value = Number(
          this.rowexport_use[ax].primary_quantity10
        );
        worksheet.getCell("AS " + [ax + 4]).numFmt = "0";
        if (this.rowexport_use[ax].pcs10 > 0) {
          this.all_result_pcs11.push({
            value: this.rowexport_use[ax].pcs10,
          });
        } else {
          this.all_result_pcs11.push({
            value: 0,
          });
        }

        if (this.rowexport_use[ax].primary_quantity10 > 0) {
          this.all_result_yard11.push({
            value: this.rowexport_use[ax].primary_quantity10,
          });
        } else {
          this.all_result_yard11.push({
            value: 0,
          });
        }
      } else if (this.rowexport_use[ax].reason11 == "NA") {
        worksheet.getCell("AR " + [ax + 4]).value = Number(
          this.rowexport_use[ax].pcs11
        );
        worksheet.getCell("AR " + [ax + 4]).numFmt = "0";
        worksheet.getCell("AS " + [ax + 4]).value = Number(
          this.rowexport_use[ax].primary_quantity11
        );
        worksheet.getCell("AS " + [ax + 4]).numFmt = "0";
        if (this.rowexport_use[ax].pcs11 > 0) {
          this.all_result_pcs11.push({
            value: this.rowexport_use[ax].pcs11,
          });
        } else {
          this.all_result_pcs11.push({
            value: 0,
          });
        }

        if (this.rowexport_use[ax].primary_quantity11 > 0) {
          this.all_result_yard11.push({
            value: this.rowexport_use[ax].primary_quantity11,
          });
        } else {
          this.all_result_yard11.push({
            value: 0,
          });
        }
      }
    },
  },
};
</script>

<style></style>
