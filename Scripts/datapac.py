from datapackage import Package

package = Package()
package.infer(r"data\avg_salary_.csv")
package.commit()
package.save(r"../datapackage.json")
